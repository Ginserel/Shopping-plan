/*
 * Minimal styled .xlsx writer (bold header, wrapped text, coloured status cells).
 * SheetJS community edition can't write cell styles, so the sheet XML is built here
 * and zipped with SheetJS's own zip writer (XLSX.CFB).
 *
 * rows: arrays of cells; a cell is a string/number or { v, s } where s is a style name below.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.SPXlsx = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  // Style names -> cellXfs index (order must match CELL_XFS below).
  const STYLE = { plain: 0, title: 1, head: 2, body: 3, bold: 4, xdock: 5, picked: 6, picking: 7, release: 8, sub: 9, time: 10, released: 11 };

  const FONTS = [
    '<font><sz val="11"/><name val="Calibri"/><family val="2"/></font>',
    '<font><b/><sz val="14"/><name val="Calibri"/><family val="2"/></font>',
    '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>',
    '<font><b/><sz val="11"/><name val="Calibri"/><family val="2"/></font>',
    '<font><sz val="11"/><color rgb="FF808080"/><name val="Calibri"/><family val="2"/></font>',
    '<font><i/><sz val="10"/><color rgb="FF595959"/><name val="Calibri"/><family val="2"/></font>',
  ];
  const fill = (rgb) => `<fill><patternFill patternType="solid"><fgColor rgb="FF${rgb}"/><bgColor indexed="64"/></patternFill></fill>`;
  const FILLS = [
    '<fill><patternFill patternType="none"/></fill>',
    '<fill><patternFill patternType="gray125"/></fill>',
    fill('1F3864'), fill('F2F2F2'), fill('C6EFCE'), fill('FFEB9C'), fill('FFC7CE'), fill('DDEBF7'),
  ];
  const BORDERS = [
    '<border><left/><right/><top/><bottom/><diagonal/></border>',
    '<border><left style="thin"><color rgb="FFBFBFBF"/></left><right style="thin"><color rgb="FFBFBFBF"/></right><top style="thin"><color rgb="FFBFBFBF"/></top><bottom style="thin"><color rgb="FFBFBFBF"/></bottom><diagonal/></border>',
  ];
  const xf = (font, fillId, border, align) => `<xf numFmtId="0" fontId="${font}" fillId="${fillId}" borderId="${border}" xfId="0"${font ? ' applyFont="1"' : ''}${fillId ? ' applyFill="1"' : ''}${border ? ' applyBorder="1"' : ''}${align ? ` applyAlignment="1"><alignment ${align}/></xf>` : '/>'}`;
  const TOP = 'vertical="top" wrapText="1"';
  const CELL_XFS = [
    xf(0, 0, 0), // plain
    xf(1, 0, 0), // title
    xf(2, 2, 1, 'vertical="center" wrapText="1"'), // head
    xf(0, 0, 1, TOP), // body
    xf(3, 0, 1, TOP), // bold
    xf(4, 3, 1, TOP), // xdock
    xf(3, 4, 1, TOP), // picked
    xf(3, 5, 1, TOP), // picking
    xf(3, 6, 1, TOP), // release
    xf(5, 0, 0), // sub
    xf(3, 0, 1, 'horizontal="center" vertical="top"'), // time
    xf(3, 7, 1, TOP), // released
  ];

  const esc = (s) => String(s)
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  function colName(i) {
    let s = '';
    for (i += 1; i > 0; i = Math.floor((i - 1) / 26)) s = String.fromCharCode(65 + ((i - 1) % 26)) + s;
    return s;
  }

  function sheetXml(spec) {
    const cols = (spec.cols || []).map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join('');
    const rows = spec.rows.map((row, r) => {
      const ht = spec.heights && spec.heights[r] ? ` ht="${spec.heights[r]}" customHeight="1"` : '';
      const cells = row.map((cell, c) => {
        if (cell == null || cell === '') return '';
        const c2 = typeof cell === 'object' ? cell : { v: cell };
        const ref = colName(c) + (r + 1);
        const s = STYLE[c2.s || 'plain'] || 0;
        if (c2.v == null || c2.v === '') return s ? `<c r="${ref}" s="${s}"/>` : '';
        if (typeof c2.v === 'number' && isFinite(c2.v)) return `<c r="${ref}" s="${s}"><v>${c2.v}</v></c>`;
        return `<c r="${ref}" s="${s}" t="inlineStr"><is><t xml:space="preserve">${esc(c2.v)}</t></is></c>`;
      }).join('');
      return `<row r="${r + 1}"${ht}>${cells}</row>`;
    }).join('');
    const freeze = spec.freezeRows
      ? `<sheetViews><sheetView workbookViewId="0"><pane ySplit="${spec.freezeRows}" topLeftCell="A${spec.freezeRows + 1}" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>`
      : '<sheetViews><sheetView workbookViewId="0"/></sheetViews>';
    const merges = (spec.merges || []).length ? `<mergeCells count="${spec.merges.length}">${spec.merges.map((m) => `<mergeCell ref="${m}"/>`).join('')}</mergeCells>` : '';
    const filter = spec.autoFilter ? `<autoFilter ref="${spec.autoFilter}"/>` : '';
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      `${freeze}<sheetFormatPr defaultRowHeight="15"/>${cols ? `<cols>${cols}</cols>` : ''}<sheetData>${rows}</sheetData>${filter}${merges}` +
      '<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>' +
      '<pageSetup orientation="landscape" fitToHeight="0"/></worksheet>';
  }

  // sheets: [{ name, cols, rows, heights, freezeRows, merges, autoFilter }] -> Uint8Array (.xlsx)
  function build(sheets, CFB) {
    CFB = CFB || (typeof XLSX !== 'undefined' && XLSX.CFB);
    if (!CFB) throw new Error('SheetJS (XLSX.CFB) is needed to write .xlsx');
    const enc = new TextEncoder();
    const names = sheets.map((s, i) => String(s.name || `Sheet${i + 1}`).replace(/[\\/?*[\]:]/g, ' ').slice(0, 31));
    const files = {
      '[Content_Types].xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('') +
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>',
      '_rels/.rels': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
      'xl/workbook.xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' +
        names.map((n, i) => `<sheet name="${esc(n)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('') + '</sheets></workbook>',
      'xl/_rels/workbook.xml.rels': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('') +
        `<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`,
      'xl/styles.xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
        `<fonts count="${FONTS.length}">${FONTS.join('')}</fonts><fills count="${FILLS.length}">${FILLS.join('')}</fills><borders count="${BORDERS.length}">${BORDERS.join('')}</borders>` +
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
        `<cellXfs count="${CELL_XFS.length}">${CELL_XFS.join('')}</cellXfs>` +
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>',
    };
    sheets.forEach((s, i) => { files[`xl/worksheets/sheet${i + 1}.xml`] = sheetXml(s); });
    const zip = CFB.utils.cfb_new();
    for (const [path, text] of Object.entries(files)) CFB.utils.cfb_add(zip, path, enc.encode(text));
    const out = CFB.write(zip, { fileType: 'zip', type: 'array', compression: true });
    return out instanceof Uint8Array ? out : new Uint8Array(out);
  }

  return { build, STYLE, _colName: colName };
});
