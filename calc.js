/*
 * Stores calculated results in the formula cells of the built workbook, so the figures show
 * as soon as the file opens (Excel's Protected View does not calculate).
 * Uses HyperFormula (GPL-3.0 licence key) to calculate. Works in the browser (window.EECalc)
 * and in Node (module.exports).
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory(typeof require === 'function' ? require('hyperformula') : null);
  } else {
    root.EECalc = factory(root.HyperFormula);
  }
})(typeof self !== 'undefined' ? self : this, function (HyperFormulaLib) {
  'use strict';

  function unescapeXml(s) {
    return s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
            .replace(/&apos;/g, "'").replace(/&amp;/g, '&');
  }
  function escapeXml(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }
  function colToNum(col) {
    var n = 0;
    for (var i = 0; i < col.length; i++) n = n * 26 + (col.charCodeAt(i) - 64);
    return n - 1;   // 0-based
  }

  /** Read every cell of one worksheet XML: {ref, row, col, f (formula text or null), v (constant or null)} */
  function readCells(xml) {
    var out = [];
    var re = /<c r="([A-Z]+)(\d+)"([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
    var m;
    while ((m = re.exec(xml)) !== null) {
      var col = colToNum(m[1]), row = +m[2] - 1, attrs = m[3], inner = m[4] || '';
      var tm = attrs.match(/\st="([^"]+)"/);
      var t = tm ? tm[1] : 'n';
      var fm = inner.match(/<f[^>]*>([\s\S]*?)<\/f>/);
      var f = fm ? unescapeXml(fm[1]) : null;
      var v = null;
      if (!f) {
        if (t === 'inlineStr') {
          var tt = inner.match(/<t[^>]*>([\s\S]*?)<\/t>/);
          v = tt ? unescapeXml(tt[1]) : '';
        } else {
          var vm = inner.match(/<v>([\s\S]*?)<\/v>/);
          if (vm) {
            var raw = unescapeXml(vm[1]);
            v = (t === 'str' || t === 's' || t === 'e') ? raw : (raw === '' ? null : +raw);
          }
        }
      }
      if (f !== null || v !== null) out.push({ ref: m[1] + (row + 1), row: row, col: col, f: f, v: v });
    }
    return out;
  }

  /**
   * sheets: array of {name, xml}. Returns {name: [{ref, type, value}]} for every formula cell.
   * type is 'n' (number), 'str', 'b' or 'e'.
   */
  function calculate(sheets) {
    var hfSheets = {};
    var cellsBySheet = {};
    var maxR = 0, maxC = 0;
    sheets.forEach(function (s) {
      var cells = readCells(s.xml);
      cellsBySheet[s.name] = cells;
      cells.forEach(function (c) { if (c.row > maxR) maxR = c.row; if (c.col > maxC) maxC = c.col; });
    });
    sheets.forEach(function (s) {
      var grid = [];
      for (var r = 0; r <= maxR; r++) {
        var row = [];
        for (var c = 0; c <= maxC; c++) row.push(null);
        grid.push(row);
      }
      cellsBySheet[s.name].forEach(function (cell) {
        grid[cell.row][cell.col] = cell.f !== null ? '=' + cell.f : (cell.v === '' ? null : cell.v);
      });
      hfSheets[s.name] = grid;
    });
    var HF = (HyperFormulaLib && HyperFormulaLib.HyperFormula) || HyperFormulaLib;
    var hf = HF.buildFromSheets(hfSheets, { licenseKey: 'gpl-v3' });
    var results = {};
    sheets.forEach(function (s) {
      var sid = hf.getSheetId(s.name);
      results[s.name] = cellsBySheet[s.name].filter(function (c) { return c.f !== null; }).map(function (c) {
        var v = hf.getCellValue({ sheet: sid, row: c.row, col: c.col });
        if (v && typeof v === 'object' && v.type !== undefined) return { ref: c.ref, type: 'e', value: v.value };
        if (v === null || v === undefined) return { ref: c.ref, type: 'str', value: '' };
        if (typeof v === 'boolean') return { ref: c.ref, type: 'b', value: v ? 1 : 0 };
        if (typeof v === 'number') return { ref: c.ref, type: 'n', value: v };
        return { ref: c.ref, type: 'str', value: String(v) };
      });
    });
    hf.destroy();
    return results;
  }

  /** Write the calculated results into one worksheet XML string. */
  function store(xml, results) {
    var byRef = {};
    results.forEach(function (r) { byRef[r.ref] = r; });
    return xml.replace(/<c r="([A-Z]+\d+)"([^>]*)><f>([\s\S]*?)<\/f>(?:<v>[^<]*<\/v>|<v\/>|<v><\/v>)?<\/c>/g,
      function (whole, ref, attrs, f) {
        var r = byRef[ref];
        if (!r) return whole;
        attrs = attrs.replace(/\st="[^"]*"/, '');
        if (r.type === 'n') return '<c r="' + ref + '"' + attrs + '><f>' + f + '</f><v>' + r.value + '</v></c>';
        if (r.type === 'b') return '<c r="' + ref + '"' + attrs + ' t="b"><f>' + f + '</f><v>' + r.value + '</v></c>';
        if (r.type === 'e') return '<c r="' + ref + '"' + attrs + ' t="e"><f>' + f + '</f><v>' + escapeXml(r.value) + '</v></c>';
        return '<c r="' + ref + '"' + attrs + ' t="str"><f>' + f + '</f><v>' + escapeXml(r.value) + '</v></c>';
      });
  }

  return { calculate: calculate, store: store, readCells: readCells };
});
