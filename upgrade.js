/*
 * Original DOCX preservation, numbering preview,
 * script-specific export fonts and footer controls.
 *
 * Load this file after the main script in index.html.
 */

let SOURCE = null;
let SOURCE_TEXT = null;

const WNS =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

const xmlRead = text =>
  new DOMParser().parseFromString(text, 'application/xml');

const xmlWrite = document =>
  new XMLSerializer().serializeToString(document);

const wx = (element, name) =>
  element.getElementsByTagNameNS(WNS, name)[0];

const attr = (element, name = 'val') =>
  element?.getAttributeNS(WNS, name) || '';

function child(element, name) {
  let result = Array.from(element.children).find(
    node => node.namespaceURI === WNS && node.localName === name
  );

  if (!result) {
    result = element.ownerDocument.createElementNS(WNS, 'w:' + name);
    element.insertBefore(result, element.firstChild);
  }

  return result;
}

function setFonts(properties, font) {
  const fonts = child(properties, 'rFonts');

  [
    'asciiTheme',
    'hAnsiTheme',
    'cstheme',
    'csTheme',
    'eastAsiaTheme'
  ].forEach(name => fonts.removeAttributeNS(WNS, name));

  ['ascii', 'hAnsi', 'cs', 'eastAsia'].forEach(name => {
    fonts.setAttributeNS(WNS, 'w:' + name, font);
  });
}

/*
 * Update fonts without rebuilding the document.
 *
 * Numbering references, starts, restarts, sections and all other package
 * parts remain in place. Symbol fonts for original bullets are preserved.
 */
function fontXML(document, numbering = false) {
  if (numbering) {
    Array.from(document.getElementsByTagNameNS(WNS, 'lvl')).forEach(level => {
      if (attr(wx(level, 'numFmt')) !== 'bullet') {
        setFonts(child(level, 'rPr'), $('nfont').value);
      }
    });
    return;
  }

  const markerEnds = new Map();
  const positions = new Map();

  for (const paragraph of Array.from(
    document.getElementsByTagNameNS(WNS, 'p')
  )) {
    const text = Array.from(
      paragraph.getElementsByTagNameNS(WNS, 't')
    ).map(node => node.textContent).join('');

    const marker = text.match(
      /^\s*(?:[0-9०-९]+(?:\.[0-9०-९]+)*[.)।]|\([0-9०-९]+\)|[क-हa-zA-Z][.)])(?=\s)/
    );

    if (marker) markerEnds.set(paragraph, marker[0].length);
  }

  for (const run of Array.from(
    document.getElementsByTagNameNS(WNS, 'r')
  )) {
    let paragraph = run.parentNode;

    while (paragraph && paragraph.localName !== 'p') {
      paragraph = paragraph.parentNode;
    }

    let position = positions.get(paragraph) || 0;

    const properties = Array.from(run.children).find(
      node => node.localName === 'rPr'
    );

    const nodes = Array.from(run.childNodes).filter(
      node => node !== properties
    );

    for (const node of nodes) {
      let parts = node.localName === 't'
        ? (
            node.textContent.match(
              /[\u0900-\u097f\u200c\u200d]+|[^\u0900-\u097f\u200c\u200d]+/gu
            ) || ['']
          )
        : [null];

      if (node.localName === 't' && markerEnds.has(paragraph)) {
        const boundary = markerEnds.get(paragraph) - position;

        if (boundary > 0 && boundary < node.textContent.length) {
          parts = [
            node.textContent.slice(0, boundary),
            ...(
              node.textContent.slice(boundary).match(
                /[\u0900-\u097f\u200c\u200d]+|[^\u0900-\u097f\u200c\u200d]+/gu
              ) || []
            )
          ];
        }
      }

      for (const text of parts) {
        const newRun = run.cloneNode(false);

        if (properties) {
          newRun.appendChild(properties.cloneNode(true));
        }

        const isMarker =
          text !== null &&
          position < (markerEnds.get(paragraph) || 0);

        const isDevanagari =
          text !== null && /[\u0900-\u097f]/u.test(text);

        const font = isMarker
          ? $('nfont').value
          : isDevanagari
            ? $('font').value
            : $('efont').value;

        setFonts(child(newRun, 'rPr'), font);

        const newNode = node.cloneNode(true);

        if (text !== null) {
          newNode.textContent = text;
          newNode.setAttributeNS(
            'http://www.w3.org/XML/1998/namespace',
            'xml:space',
            'preserve'
          );
        }

        newRun.appendChild(newNode);
        run.parentNode.insertBefore(newRun, run);

        if (text !== null) position += text.length;
      }
    }

    positions.set(paragraph, position);
    run.remove();
  }
}

async function finishDocx(input) {
  const bytes = input instanceof Blob
    ? await input.arrayBuffer()
    : input;

  const zip = await JSZip.loadAsync(bytes);

  const parts = Object.keys(zip.files).filter(path =>
    /^word\/(document|header\d*|footer\d*|footnotes|endnotes|numbering)\.xml$/.test(path)
  );

  for (const path of parts) {
    const document = xmlRead(await zip.file(path).async('string'));

    if (document.getElementsByTagName('parsererror').length) {
      throw new Error('Could not read the Word XML part: ' + path);
    }

    fontXML(document, path === 'word/numbering.xml');

    if (
      path === 'word/document.xml' &&
      $('footerMode').value === 'none'
    ) {
      Array.from(
        document.getElementsByTagNameNS(WNS, 'footerReference')
      ).forEach(node => node.remove());
    }

    zip.file(path, xmlWrite(document));
  }

  return zip.generateAsync({
    type: 'blob',
    mimeType:
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
  });
}

const baseBuild = build;

build = async function (output, name) {
  if ($('mode').value === 'preserve' && SOURCE) {
    return finishDocx(SOURCE);
  }

  return finishDocx(await baseBuild(output, name));
};

const baseRun = run;

run = function () {
  const preserve = $('mode').value === 'preserve';
  const originalDocx = preserve && !!SOURCE;

  if (originalDocx && SOURCE_TEXT !== null) {
    FM = SOURCE_TEXT;
    $('in').value = pl(FM);
  }

  $('in').readOnly = originalDocx;
  $('digits').disabled = preserve;
  $('size').disabled = originalDocx;

  $('modeHelp').textContent = preserve
    ? originalDocx
      ? 'Original DOCX mode: list definitions, restarts, tables, section settings and existing footer are retained. Font choices and footer removal apply on download. Changing fonts may change line wrapping. Switch mode to edit/restructure text.'
      : 'Original text mode: typed numbering, punctuation, order and blank lines stay unchanged. No new numbering is added. For exact Word list preservation, upload a .docx file.'
    : 'New numbering replaces typed values. Review the preview, especially detected headings and levels. तपसिल restarts at १ except in Irrespective mode.';

  if (!$('in').value.trim()) {
    ops = null;
    $('out').innerHTML = '';
    $('msg').textContent = '';
    $('dl').disabled = $('dld').disabled = $('rt').disabled = true;
    return;
  }

  if (!preserve) {
    $('dld').title = '';
    baseRun();
    return;
  }

  $('msg').textContent = '';

  const text = FM && pl(FM) === $('in').value
    ? FM
    : $('in').value;

  ops = text
    .replace(/\r\n/g, '\n')
    .split('\n')
    .map(line => ({ k: 'l', t: line }));

  preview(ops);
  $('out').style.whiteSpace = 'pre-wrap';

  $('out').querySelectorAll('div').forEach(element => {
    if (!element.textContent) element.innerHTML = '<br>';
  });

  const note = $('out').lastElementChild;

  if (note?.classList.contains('note')) {
    note.textContent = originalDocx
      ? 'Footer: retained from original DOCX (unless removed).'
      : 'Footer: page number (unless removed).';
  }

  $('dl').disabled = false;

  /*
   * RTF cannot faithfully retain arbitrary original DOCX numbering,
   * layout and section structure. Use original DOCX export instead.
   */
  $('dld').disabled = originalDocx;
  $('dld').title = originalDocx
    ? 'Use .docx to retain original Word numbering and layout.'
    : '';

  $('rt').disabled = originalDocx;
};

const basePreview = preview;

preview = function (output) {
  $('out').style.whiteSpace = 'normal';
  basePreview(output);

  if ($('footerMode').value === 'none') {
    const note = $('out').lastElementChild;
    if (note?.classList.contains('note')) {
      const separator = note.previousElementSibling;
      note.remove();
      if (separator?.tagName === 'HR') separator.remove();
    }
  }
};

for (const id of [
  'mode',
  'font',
  'efont',
  'nfont',
  'size',
  'digits',
  'footerMode'
]) {
  $(id).onchange = () => run();
}

$('go').onclick = run;

$('in').addEventListener('input', () => {
  ops = null;
  $('dl').disabled = $('dld').disabled = $('rt').disabled = true;
  $('msg').textContent =
    'Input changed. Press Structure to refresh the output before downloading.';
});

$('dl').onclick = async () => {
  if (!ops) return;

  const button = $('dl');
  button.disabled = true;

  try {
    save(await build(ops), outName(ops, 'docx'));
  } catch (error) {
    $('msg').textContent = 'Download failed: ' + error.message;
  } finally {
    button.disabled = !ops;
  }
};

/*
 * Compatibility RTF export:
 * - use Devanagari numbering when selected;
 * - explicitly apply the Devanagari font to Unicode text;
 * - optionally remove the footer.
 */
const baseRtf = mkRtf;

mkRtf = function (output) {
  let text = baseRtf(output);

  if ($('digits').value === 'ne') {
    text = text.replace(
      /\\levelnfc0\\levelnfcn0/g,
      '\\levelnfc48\\levelnfcn48'
    );
  }

  text = text.replace(
    /(?:\\u-?\d+\?)+/g,
    chunk => '{\\f0 ' + chunk + '}'
  );

  if ($('footerMode').value === 'none') {
    const start = text.indexOf('{\\footer');

    if (start >= 0) {
      let depth = 0;
      let end = start;

      for (; end < text.length; end++) {
        if (text[end] === '{' && text[end - 1] !== '\\') depth++;

        if (
          text[end] === '}' &&
          text[end - 1] !== '\\' &&
          --depth === 0
        ) {
          end++;
          break;
        }
      }

      text = text.slice(0, start) + text.slice(end);
    }
  }

  return text;
};

/*
 * Materialize Word list labels for the text preview and for restructuring.
 * Original-mode downloads use the original DOCX package and definitions.
 */
const baseDocxM = docxM;

docxM = async function (arrayBuffer) {
  const text = await baseDocxM(arrayBuffer);
  const zip = await JSZip.loadAsync(arrayBuffer);

  if (!zip.file('word/numbering.xml')) {
    SOURCE_TEXT = text;
    return text;
  }

  const numbering = xmlRead(
    await zip.file('word/numbering.xml').async('string')
  );

  const document = xmlRead(
    await zip.file('word/document.xml').async('string')
  );

  const stylesDocument = zip.file('word/styles.xml')
    ? xmlRead(await zip.file('word/styles.xml').async('string'))
    : null;

  const styles = new Map(
    stylesDocument
      ? Array.from(
          stylesDocument.getElementsByTagNameNS(WNS, 'style')
        ).map(style => [attr(style, 'styleId'), style])
      : []
  );

  const numbers = new Map(
    Array.from(
      numbering.getElementsByTagNameNS(WNS, 'num')
    ).map(number => [attr(number, 'numId'), number])
  );

  const abstracts = new Map(
    Array.from(
      numbering.getElementsByTagNameNS(WNS, 'abstractNum')
    ).map(number => [attr(number, 'abstractNumId'), number])
  );

  const counters = {};
  const lines = text.split('\n');
  let offset = 0;

  function numberingProperties(paragraph) {
    const properties = wx(paragraph, 'pPr');
    const numberProperties = properties && wx(properties, 'numPr');

    let id = attr(numberProperties && wx(numberProperties, 'numId'));
    let level = attr(numberProperties && wx(numberProperties, 'ilvl'));
    let styleId = attr(properties && wx(properties, 'pStyle'));

    const seen = new Set();

    while (styleId && !seen.has(styleId)) {
      seen.add(styleId);

      const style = styles.get(styleId);
      if (!style) break;

      const inherited = wx(style, 'numPr');

      id = id || attr(inherited && wx(inherited, 'numId'));
      level = level || attr(inherited && wx(inherited, 'ilvl'));
      styleId = attr(wx(style, 'basedOn'));
    }

    return { id, level: +(level || 0) };
  }

  function roman(number) {
    let result = '';

    const symbols = [
      [1000, 'M'],
      [900, 'CM'],
      [500, 'D'],
      [400, 'CD'],
      [100, 'C'],
      [90, 'XC'],
      [50, 'L'],
      [40, 'XL'],
      [10, 'X'],
      [9, 'IX'],
      [5, 'V'],
      [4, 'IV'],
      [1, 'I']
    ];

    for (const [value, symbol] of symbols) {
      while (number >= value) {
        result += symbol;
        number -= value;
      }
    }

    return result;
  }

  function formatNumber(number, format) {
    if (format === 'hindiNumbers') return dev(number);
    if (format === 'lowerRoman') return roman(number).toLowerCase();
    if (format === 'upperRoman') return roman(number);

    if (/Letter$/.test(format)) {
      let result = '';

      for (; number > 0; number = Math.floor((number - 1) / 26)) {
        result =
          String.fromCharCode(65 + (number - 1) % 26) + result;
      }

      return format === 'lowerLetter'
        ? result.toLowerCase()
        : result;
    }

    return String(number);
  }

  for (const paragraph of Array.from(
    document.getElementsByTagNameNS(WNS, 'p')
  )) {
    const { id, level } = numberingProperties(paragraph);
    const number = numbers.get(id);

    const abstract = number &&
      abstracts.get(attr(wx(number, 'abstractNumId')));

    if (abstract && id !== '0') {
      const levels = new Map(
        Array.from(abstract.children)
          .filter(element => element.localName === 'lvl')
          .map(element => [+attr(element, 'ilvl'), element])
      );

      const overrides = new Map(
        Array.from(number.children)
          .filter(element => element.localName === 'lvlOverride')
          .map(element => [+attr(element, 'ilvl'), element])
      );

      const definition = key =>
        (overrides.get(key) && wx(overrides.get(key), 'lvl')) ||
        levels.get(key);

      const start = key => +(
        attr(overrides.get(key) && wx(overrides.get(key), 'startOverride')) ||
        attr(definition(key) && wx(definition(key), 'start')) ||
        1
      );

      const counts = counters[id] || (counters[id] = []);
      const current = definition(level);

      if (current) {
        counts[level] = counts[level] === undefined
          ? start(level)
          : counts[level] + 1;

        for (let key = level + 1; key < 9; key++) {
          const restart = attr(
            definition(key) && wx(definition(key), 'lvlRestart')
          );

          if (
            restart !== '0' &&
            (restart === '' || level < +restart)
          ) {
            counts[key] = undefined;
          }
        }

        const pattern = attr(wx(current, 'lvlText'));
        const format = attr(wx(current, 'numFmt'));

        const label = format === 'bullet'
          ? pattern
          : pattern.replace(/%([1-9])/g, (_, value) => {
              const key = +value - 1;
              return formatNumber(
                counts[key] ?? start(key),
                attr(definition(key) && wx(definition(key), 'numFmt'))
              );
            });

        if (lines[offset] !== undefined) {
          lines[offset] = label + '\t' + lines[offset];
        }
      }
    }

    offset +=
      1 +
      paragraph.getElementsByTagNameNS(WNS, 'br').length +
      paragraph.getElementsByTagNameNS(WNS, 'cr').length;
  }

  SOURCE_TEXT = lines.join('\n');
  return SOURCE_TEXT;
};

run();
