/*
 * EE Workbook Builder - core logic (no UI code).
 * Works in the browser (window.EECore) and in Node (module.exports).
 *
 * Responsibilities
 *  1. Find the header row of an employee list and map its columns by name.
 *  2. Normalise every employee value to the codes the workbook formulas expect.
 *  3. Record every value that could not be read as an "issue" (never silently drop data).
 *  4. Write the normalised rows and settings into the template workbook XML, keeping
 *     all formulas, formatting, freeze panes and drop-downs untouched.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory(typeof require === 'function' ? require('xlsx-js-style') : null,
                             typeof require === 'function' ? require('jszip') : null,
                             typeof require === 'function' ? require('./calc.js') : null);
  } else {
    root.EECore = factory(root.XLSX, root.JSZip, root.EECalc);
  }
})(typeof self !== 'undefined' ? self : this, function (XLSX, JSZip, EECalc) {
  'use strict';

  // ---------------------------------------------------------------- constants
  var FIRST_ROW = 10;          // first employee row in "EEA2 & EEA4 INFO"
  var LAST_ROW = 6009;         // last row with formulas in the template (6,000 employees)
  var DATA_SHEET = 'EEA2 & EEA4 INFO';

  var LEVEL_TEXT = {
    1: 'Top Management', 2: 'Senior Management', 3: 'Professionally Qualified',
    4: 'Skilled Technical', 5: 'Semi-Skilled', 6: 'Unskilled'
  };
  var LEVEL_LOOKUP = {
    'top management': 1, 'senior management': 2,
    'professionally qualified': 3, 'professionally qualified and experienced specialists and mid management': 3,
    'skilled technical': 4, 'skilled technical and academically qualified workers': 4,
    'skilled': 4, 'skilled junior management': 4, 'skilled technical and junior management': 4,
    'semi skilled': 5, 'semi skilled and discretionary decision making': 5,
    'unskilled': 6, 'unskilled and defined decision making': 6,
    '1': 1, '2': 2, '3': 3, '4': 4, '5': 5, '6': 6
  };
  var RACE_LOOKUP = {
    'african': 'AFRICAN', 'a': 'AFRICAN', 'black african': 'AFRICAN',
    'coloured': 'COLOURED', 'c': 'COLOURED',
    'indian': 'INDIAN', 'i': 'INDIAN',
    'white': 'WHITE', 'w': 'WHITE'
  };
  var RACE_WORDS = ['african', 'coloured', 'indian', 'white', 'asian', 'black'];
  var SA_WORDS = ['rsa', 'sa', 'za', 'zaf', 'south africa', 'south african'];
  var BLANK_WORDS = ['', 'unspecified', 'unknown', 'n/a', '0'];

  // Header synonyms. Keys are the internal field names.
  var HEADER_SYNONYMS = {
    code: ['empl code', 'emp number', 'employee code', 'employee number', 'emp code', 'empl no', 'emp no', 'employee no', 'emp id'],
    surname: ['surname', 'emp surname', 'employee surname', 'last name', 'lastname'],
    names: ['full names', 'emp name', 'employee name', 'first names', 'names', 'full name', 'first name', 'firstname'],
    gender: ['gender', 'sex'],
    race: ['race', 'population group', 'racial group'],
    engaged: ['date engaged', 'contract start date', 'start date', 'date appointed', 'engagement date', 'appointment date', 'dateengaged'],
    term: ['termination date', 'term date', 'date terminated', 'terminationdate'],
    prev_pos: ['previous position'],
    cur_pos: ['current position', 'job title', 'position'],
    level: ['occ levels if available', 'occ levels', 'occupational level', 'occ level', 'level'],
    level_no: ['level no', 'level number'],
    level_start: ['occ level at start of period', 'occ level at start of period for promotions', 'level at start of period', 'start level'],
    perm: ['perm temp', 'perm/temp', 'status perm contract', 'status', 'employment type'],
    nationality: ['nationality', 'citizenship', 'nationality country code'],
    fn_flag: ['fn', 'foreign national', 'foreign'],
    total_rem: ['total', 'total remuneration', 'total pay'],
    disabled: ['disabled', 'disability']
  };
  var REQUIRED_FOR_HEADER = ['code', 'surname', 'names', 'gender', 'race', 'engaged', 'level', 'perm', 'nationality', 'disabled'];
  var FATAL_IF_MISSING = ['code', 'gender', 'race', 'engaged', 'level'];

  // ---------------------------------------------------------------- helpers
  function norm(s) {
    return String(s == null ? '' : s).replace(/\s+/g, ' ').replace(/[^A-Za-z0-9]+/g, ' ').trim().toLowerCase();
  }
  function esc(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }
  function colToNum(col) {
    var n = 0;
    for (var i = 0; i < col.length; i++) n = n * 26 + (col.charCodeAt(i) - 64);
    return n;
  }
  function numToCol(n) {
    var s = '';
    while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }
  function daysIn(y, m) { return new Date(Date.UTC(y, m, 0)).getUTCDate(); }

  // Excel serial number for a calendar date (1900 date system).
  function serialFromYMD(y, m, d) {
    return Date.UTC(y, m - 1, d) / 86400000 + 25569;
  }
  function ymdFromSerial(serial) {
    var ms = Math.round((serial - 25569) * 86400000);
    var dt = new Date(ms);
    return { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate() };
  }

  /**
   * Parse a date cell. Returns {serial, iso, problem}.
   *  - serial: Excel serial number or null
   *  - problem: '' when fine, otherwise a short description
   * Accepts Excel date numbers, dd/mm/yyyy text (1/5/2024, 31/082026, 31-08-2026),
   * yyyy-mm-dd text. Rejects impossible dates such as 31/02/2025 and time-only values.
   */
  function parseDate(v, order) {
    order = order || 'dmy';
    if (v === null || v === undefined || v === '') return { serial: null, iso: '', problem: '' };
    var y, m, d;
    if (typeof v === 'number') {
      if (v === 0) return { serial: null, iso: '', problem: 'value is 0' };
      if (v < 1) return { serial: null, iso: '', problem: 'time only, no date' };
      var ymd = ymdFromSerial(v);
      y = ymd.y; m = ymd.m; d = ymd.d;
    } else if (v instanceof Date) {
      y = v.getUTCFullYear(); m = v.getUTCMonth() + 1; d = v.getUTCDate();
    } else {
      var s = String(v).trim();
      if (s === '' || s === '0') return { serial: null, iso: '', problem: s === '0' ? 'value is 0' : '' };
      var iso = s.match(/^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})/);
      if (iso) { y = +iso[1]; m = +iso[2]; d = +iso[3]; }
      else {
        var parts = s.split(/[^0-9]+/).filter(function (x) { return x.length > 0; });
        if (parts.length === 3) {
          if (order === 'mdy') { m = +parts[0]; d = +parts[1]; y = +parts[2]; }
          else if (order === 'ymd') { y = +parts[0]; m = +parts[1]; d = +parts[2]; }
          else { d = +parts[0]; m = +parts[1]; y = +parts[2]; }
        }
        else if (parts.length === 2 && parts[1].length === 6) {
          if (order === 'mdy') { m = +parts[0]; d = +parts[1].slice(0, 2); }
          else { d = +parts[0]; m = +parts[1].slice(0, 2); }
          y = +parts[1].slice(2);
        }
        else if (parts.length === 1 && parts[0].length === 8) { d = +parts[0].slice(0, 2); m = +parts[0].slice(2, 4); y = +parts[0].slice(4); }
        else return { serial: null, iso: '', problem: 'not a readable date "' + s + '"' };
      }
    }
    if (!(y >= 1940 && y <= 2100)) return { serial: null, iso: '', problem: 'year ' + y + ' not plausible' };
    if (!(m >= 1 && m <= 12) || !(d >= 1 && d <= daysIn(y, m))) {
      return { serial: null, iso: '', problem: 'impossible date ' + d + '/' + m + '/' + y };
    }
    var pad = function (n) { return (n < 10 ? '0' : '') + n; };
    return { serial: serialFromYMD(y, m, d), iso: y + '-' + pad(m) + '-' + pad(d), problem: '' };
  }

  function parseNumber(v) {
    if (typeof v === 'number') return { value: v, problem: '' };
    if (v === null || v === undefined || String(v).trim() === '') return { value: 0, problem: '' };
    var s = String(v).replace(/[R\s,]/g, '');
    var n = Number(s);
    if (isFinite(n)) return { value: n, problem: '' };
    return { value: 0, problem: 'not a number "' + v + '"' };
  }

  // ---------------------------------------------------------------- reading
  /** Read the first sheet (or a named sheet) of an uploaded file into a 2-D array of raw values. */
  var cachedBuf = null, cachedWb = null;
  /** Parse an uploaded file once; repeated calls with the same buffer reuse the result. */
  function openWorkbook(arrayBuffer) {
    if (arrayBuffer !== cachedBuf) {
      cachedWb = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array', cellDates: false, cellNF: false, cellText: false });
      cachedBuf = arrayBuffer;
    }
    return cachedWb;
  }

  function readRows(arrayBuffer, sheetName) {
    var wb = openWorkbook(arrayBuffer);
    var name = sheetName || wb.SheetNames[0];
    var ws = wb.Sheets[name];
    if (!ws) throw new Error('Sheet "' + name + '" not found');
    var firstRow = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']).s.r : 0;
    var rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '', blankrows: true });
    rows.firstRow = firstRow;   // 0-based Excel row of rows[0]
    return { sheetNames: wb.SheetNames, sheetName: name, rows: rows };
  }

  /** Locate the header row and map internal field names to column indexes. */
  function detectHeader(rows) {
    var best = null;
    var limit = Math.min(rows.length, 60);
    for (var r = 0; r < limit; r++) {
      var cells = rows[r].map(norm);
      var map = {};
      var hits = 0;
      for (var f in HEADER_SYNONYMS) {
        var syn = HEADER_SYNONYMS[f].map(function (x) { return x.replace(/ /g, ''); });
        for (var c = 0; c < cells.length; c++) {
          if (syn.indexOf(cells[c].replace(/ /g, '')) !== -1 && map[f] === undefined) { map[f] = c; hits++; break; }
        }
      }
      var req = REQUIRED_FOR_HEADER.filter(function (k) { return map[k] !== undefined; }).length;
      if (req >= 5 && (!best || req > best.req)) best = { row: r, map: map, req: req, cells: cells };
    }
    if (!best) throw new Error('No header row found. The file needs columns such as Empl Code, Gender, RACE, Date Engaged, Occ Levels.');

    // Remuneration blocks by position between the known header names
    var cells = best.cells, fixedStart = -1, fixedEnd = -1, varStart = -1, varEnd = -1;
    for (var c2 = 0; c2 < cells.length; c2++) {
      var h = cells[c2];
      if (fixedStart < 0 && h.indexOf('salary') === 0) fixedStart = c2;
      if (h.indexOf('other fixed') === 0) fixedEnd = c2;
      if (varStart < 0 && h.indexOf('bonus') === 0) varStart = c2;
      if (h.indexOf('other variable') === 0) varEnd = c2;
    }
    var fixedCols = fixedStart >= 0 && fixedEnd >= fixedStart ? range(fixedStart, fixedEnd) : [];
    var varCols = varStart >= 0 && varEnd >= varStart ? range(varStart, varEnd) : [];
    var missing = REQUIRED_FOR_HEADER.filter(function (k) { return best.map[k] === undefined; });
    var fatal = missing.filter(function (k) { return FATAL_IF_MISSING.indexOf(k) !== -1; });
    return { headerRow: best.row, map: best.map, fixedCols: fixedCols, varCols: varCols,
             missing: missing, fatal: fatal };
  }
  function range(a, b) { var out = []; for (var i = a; i <= b; i++) out.push(i); return out; }

  // ---------------------------------------------------------------- normalising
  function cleanText(v) {
    if (v === null || v === undefined) return '';
    if (typeof v === 'number') return v === 0 ? '' : String(v);
    var t = String(v).replace(/\s+/g, ' ').trim();
    return /^(none|null|n\/a|#n\/a)$/i.test(t) ? '' : t;
  }

  /**
   * Normalise every employee row.
   * Returns {employees: [...], issues: [...], rowsRead, skipped}
   * Each employee: {src (source row number), code, surname, names, gender, race,
   *   engaged (serial|null), term (serial|null), prevPos, curPos, level (text to write),
   *   levelNo (1-6|null), levelStart (text), perm, nationality, disabled, fixed[12], variable[7]}
   */
  /**
   * Work out whether text dates such as 01/05/2024 are day-first or month-first, from the values
   * in the file (any first number above 12 means day-first; any second number above 12 means month-first).
   */
  function detectDateOrder(rows, detected) {
    var idxs = [detected.map.engaged, detected.map.term].filter(function (x) { return x !== undefined; });
    var first12 = 0, second12 = 0, slash = 0;
    rows.slice(detected.headerRow + 1).forEach(function (row) {
      idxs.forEach(function (ci) {
        var v = row[ci];
        if (typeof v !== 'string') return;
        var m = v.trim().match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/);
        if (!m) return;
        slash++;
        if (+m[1] > 12) first12++;
        if (+m[2] > 12) second12++;
      });
    });
    var order = 'dmy', conflict = false;
    if (second12 > 0 && first12 === 0) order = 'mdy';
    if (first12 > 0 && second12 > 0) conflict = true;
    return { order: order, conflict: conflict, ambiguous: slash > 0 && first12 === 0 && second12 === 0, textDates: slash };
  }

  function normalise(rows, detected, permMap, opts) {
    var dateChoice = (opts && opts.dateFormat) || 'auto';
    var dateInfo = null;
    var dateOrder = 'dmy';
    if (dateChoice === 'auto') { dateInfo = detectDateOrder(rows, detected); dateOrder = dateInfo.order; }
    else dateOrder = dateChoice;
    var noRemunRows = 0, totalOnlyRows = 0;
    var m = detected.map;
    var employees = [], issues = [], skipped = 0, noteFnBlank = false, excluded = [];
    function get(row, f) { return m[f] === undefined ? '' : row[m[f]]; }
    function issue(rec, field, problem, action) {
      issues.push({ src: rec.src, code: rec.code, name: (rec.surname + ' ' + rec.names).trim(), field: field, problem: problem, action: action || '' });
    }

    for (var r = detected.headerRow + 1; r < rows.length; r++) {
      var row = rows[r];
      var code = cleanText(get(row, 'code'));
      var surname = cleanText(get(row, 'surname'));
      var names = cleanText(get(row, 'names'));
      if (!code && !surname && !names) { skipped++; continue; }

      var rec = { src: r + (rows.firstRow || 0) + 1, code: code, surname: surname, names: names };
      var lvlText = norm(get(row, 'level'));
      if (lvlText.indexOf('independent contractor') !== -1) {
        excluded.push({ src: rec.src, code: code, name: (surname + ' ' + names).trim(), field: 'Left out',
                        problem: 'independent contractor (not an employee under the EEA), left out of the workbook',
                        action: 'None. Change the level only if this person is an employee' });
        continue;
      }

      // Gender
      var g = norm(get(row, 'gender'));
      rec.gender = (g === 'm' || g === 'male') ? 'M' : (g === 'f' || g === 'female') ? 'F' : '';
      if (!rec.gender) issue(rec, 'Gender', 'missing or unreadable (found "' + cleanText(get(row, 'gender')) + '")', 'Enter Male or Female');

      // Race
      var rc = norm(get(row, 'race'));
      rec.race = RACE_LOOKUP[rc] || '';
      if (!rec.race) issue(rec, 'Race', rc === '' || rc === '0' || rc === 'unspecified' ? 'unspecified' : 'unknown value "' + cleanText(get(row, 'race')) + '"',
                           'Enter African, Coloured, Indian or White');

      // Dates
      var eng = parseDate(get(row, 'engaged'), dateOrder);
      rec.engaged = eng.serial;
      if (eng.problem || eng.serial === null) issue(rec, 'Date engaged', eng.problem || 'missing', 'Enter a valid date (dd/mm/yyyy)');
      var trm = parseDate(get(row, 'term'), dateOrder);
      rec.term = trm.serial;
      if (trm.problem && trm.problem !== 'value is 0') issue(rec, 'Termination date', trm.problem, 'Enter a valid date or leave blank');

      // Positions
      rec.prevPos = cleanText(get(row, 'prev_pos'));
      rec.curPos = cleanText(get(row, 'cur_pos'));

      // Occupational level
      var lvlRaw = get(row, 'level');
      var lvlNoRaw = m.level_no !== undefined ? get(row, 'level_no') : '';
      var lvlKey = norm(lvlRaw);
      var lvlNo = LEVEL_LOOKUP[lvlKey];
      if (lvlNo === undefined && typeof lvlRaw === 'number' && LEVEL_LOOKUP[String(lvlRaw)]) lvlNo = LEVEL_LOOKUP[String(lvlRaw)];
      if (lvlNo === undefined && lvlNoRaw !== '' && LEVEL_LOOKUP[norm(lvlNoRaw)]) lvlNo = LEVEL_LOOKUP[norm(lvlNoRaw)];
      rec.levelNo = lvlNo || null;
      rec.level = lvlNo ? LEVEL_TEXT[lvlNo] : cleanText(lvlRaw);
      if (!lvlNo) issue(rec, 'Occupational level', lvlKey === '' || lvlKey === 'unspecified' ? 'unspecified' : 'unknown value "' + cleanText(lvlRaw) + '"',
                        'Enter the occupational level (leave open only if you do not know)');
      var lvlStartRaw = m.level_start !== undefined ? get(row, 'level_start') : '';
      var lvlStartKey = norm(lvlStartRaw);
      rec.levelStart = LEVEL_LOOKUP[lvlStartKey] ? LEVEL_TEXT[LEVEL_LOOKUP[lvlStartKey]] : cleanText(lvlStartRaw);

      // Permanent / temporary
      var p = norm(get(row, 'perm'));
      if (p === 'perm' || p === 'permanent') rec.perm = 'Permanent';
      else if (p === 'temp' || p === 'temporary' || p === 'contract' || p === 'fixed term' || p === 'casual') rec.perm = 'Temporary';
      else if (p === '' && permMap && permMap[code.toUpperCase()]) { rec.perm = permMap[code.toUpperCase()]; rec.permFromPayroll = true; }
      else { rec.perm = ''; issue(rec, 'Perm/Temp', p === '' || p === 'unspecified' ? 'unspecified' : 'unknown value "' + cleanText(get(row, 'perm')) + '"',
                                  'Enter Permanent or Temporary'); }

      // Nationality
      var n = norm(get(row, 'nationality'));
      var fnFlag = m.fn_flag !== undefined ? cleanText(get(row, 'fn_flag')) : '';
      if (m.nationality === undefined && m.fn_flag !== undefined) {
        rec.nationality = fnFlag !== '' ? 'Foreign' : 'RSA';
        if (fnFlag === '') noteFnBlank = true;
      } else if (m.nationality === undefined) {
        rec.nationality = '';
        issue(rec, 'Nationality', 'no nationality column in this file', 'Enter RSA or the foreign nationality');
      } else if (SA_WORDS.indexOf(n) !== -1) rec.nationality = 'RSA';
      else if (BLANK_WORDS.indexOf(n) !== -1) {
        rec.nationality = '';
        issue(rec, 'Nationality', 'unspecified', 'Enter RSA or the foreign nationality');
      } else if (RACE_WORDS.indexOf(n) !== -1) {
        rec.nationality = '';
        issue(rec, 'Nationality', 'holds a race value "' + cleanText(get(row, 'nationality')) + '" (columns may be shifted)', 'Enter RSA or the foreign nationality');
      } else {
        rec.nationality = cleanText(get(row, 'nationality'));
      }

      // Disability
      var dis = norm(get(row, 'disabled'));
      rec.disabled = (dis === 'yes' || dis === 'y' || dis === 'disabled' || dis === 'disability') ? 'Yes' : 'No';

      // Remuneration
      var fixed = [], variable = [];
      detected.fixedCols.forEach(function (ci) { fixed.push(parseNumber(row[ci])); });
      detected.varCols.forEach(function (ci) { variable.push(parseNumber(row[ci])); });
      rec.fixed = [], rec.variable = [];
      for (var i = 0; i < 12; i++) rec.fixed.push(fixed[i] ? fixed[i].value : 0);
      for (var j = 0; j < 7; j++) rec.variable.push(variable[j] ? variable[j].value : 0);
      fixed.concat(variable).forEach(function (x) { if (x.problem) issue(rec, 'Remuneration', x.problem, 'Enter a number (or 0)'); });
      if (detected.fixedCols.length === 0) {
        if (m.total_rem !== undefined) {
          var tot = parseNumber(get(row, 'total_rem'));
          rec.fixed[0] = tot.value;
          totalOnlyRows++;
        } else {
          noRemunRows++;
        }
      }

      if (!code) issue(rec, 'Empl code', 'missing', 'Enter a unique employee code');
      employees.push(rec);
    }
    var notes = [];
    if (noteFnBlank) notes.push('Nationality taken from the FN column: a blank FN was treated as RSA (South African).');
    var orderWords = { dmy: 'day/month/year', mdy: 'month/day/year', ymd: 'year/month/day' };
    if (dateChoice === 'auto' && dateInfo && dateInfo.textDates) {
      notes.push('Text dates read as ' + orderWords[dateOrder] + ' (detected from the file' +
        (dateInfo.conflict ? '; some dates conflict, so check the Import Log' : (dateInfo.ambiguous ? '; no date in the file settles the order, so day/month/year was used' : '')) + ').');
    } else if (dateChoice !== 'auto') {
      notes.push('Text dates read as ' + orderWords[dateOrder] + ' (chosen on the page).');
    }
    if (employees.some(function (e) { return e.permFromPayroll; })) notes.push('__PERM_NOTE__');
    if (noRemunRows) notes.push('This file has no remuneration columns, so the EEA4 remuneration figures will be blank.');
    if (totalOnlyRows) notes.push('Only a total remuneration is given for ' + totalOnlyRows + ' employee(s); it is placed in Salary/Wage (no fixed/variable split).');
    return { employees: employees, issues: issues, excluded: excluded, notes: notes, dateOrder: dateOrder, rowsRead: rows.length - detected.headerRow - 1, skipped: skipped };
  }

  /**
   * Read Perm/Temp from a payroll sheet in the same file (a sheet with Employee Code and Perm/Temp columns).
   * Returns {map: {CODE: 'Permanent'|'Temporary'}, sheet} or null when no such sheet exists.
   */
  function readPermMap(buf, employeeSheet) {
    var wb = openWorkbook(buf);
    for (var i = 0; i < wb.SheetNames.length; i++) {
      var name = wb.SheetNames[i];
      if (name === employeeSheet) continue;
      var rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: '', blankrows: true });
      var hr = -1, cc = -1, pc = -1;
      for (var r = 0; r < Math.min(rows.length, 15) && hr < 0; r++) {
        var cells = rows[r].map(function (x) { return String(x).toLowerCase().replace(/\s+/g, ''); });
        var a = cells.indexOf('employeecode'), b = cells.indexOf('perm/temp');
        if (a < 0) a = cells.indexOf('employeecode');
        if (b < 0) b = cells.indexOf('permtemp');
        if (a >= 0 && b >= 0) { hr = r; cc = a; pc = b; }
      }
      if (hr < 0) continue;
      var map = {};
      for (var k = hr + 1; k < rows.length; k++) {
        var code = cleanText(rows[k][cc]).toUpperCase();
        var pt = norm(rows[k][pc]);
        if (!code || map[code]) continue;
        if (pt === 'perm' || pt === 'permanent') map[code] = 'Permanent';
        else if (pt === 'temp' || pt === 'temporary') map[code] = 'Temporary';
      }
      return { map: map, sheet: name };
    }
    return null;
  }

  /**
   * Apply the reporting period. Employees engaged after the period end, or who left before the
   * period started, are left out of the workbook and listed on the Import Log (field 'Left out').
   */
  function applyPeriod(res, periodStart, periodEnd) {
    var keep = [], left = [], keepSrc = {};
    res.employees.forEach(function (e) {
      var reason = '';
      if (e.engaged !== null && e.engaged > periodEnd) reason = 'engaged after the period end (' + isoOf(e.engaged) + ')';
      else if (e.term !== null && e.term < periodStart) reason = 'left before the period started (' + isoOf(e.term) + ')';
      if (reason) left.push({ src: e.src, code: e.code, name: (e.surname + ' ' + e.names).trim(), field: 'Left out',
                              problem: 'not employed during the period: ' + reason, action: 'None. Check the dates if this person should be included' });
      else { keep.push(e); keepSrc[e.src] = true; }
    });
    var issues = res.issues.filter(function (i) { return keepSrc[i.src] || i.field === 'Left out'; });
    // the payroll note counts only the employees who are in the workbook
    var nPayroll = keep.filter(function (e) { return e.permFromPayroll; }).length;
    var notes = (res.notes || []).filter(function (n) { return n !== '__PERM_NOTE__'; });
    if (nPayroll) notes.unshift('Perm/Temp taken from the payroll sheet for ' + nPayroll + ' employee(s), which has no Perm/Temp column.');
    return { employees: keep, issues: issues, excluded: (res.excluded || []).concat(left), notes: notes,
             rowsRead: res.rowsRead, skipped: res.skipped };
  }
  function isoOf(serial) {
    var d = new Date(Math.round((serial - 25569) * 86400000));
    var pad = function (n) { return (n < 10 ? '0' : '') + n; };
    return d.getUTCFullYear() + '-' + pad(d.getUTCMonth() + 1) + '-' + pad(d.getUTCDate());
  }

  // Duplicate codes are a data problem the workbook also checks, so report them here too.
  function addDuplicateIssues(result) {
    var seen = {};
    result.employees.forEach(function (e) {
      if (!e.code) return;
      if (seen[e.code]) seen[e.code].push(e); else seen[e.code] = [e];
    });
    Object.keys(seen).forEach(function (k) {
      if (seen[k].length > 1) seen[k].forEach(function (e) {
        result.issues.push({ src: e.src, code: e.code, name: (e.surname + ' ' + e.names).trim(), field: 'Empl code',
                             problem: 'duplicate code (used ' + seen[k].length + ' times)', action: 'Make each code unique' });
      });
    });
    return result;
  }

  /**
   * Summary for the preview screen. snapshot is a Date-like serial (Excel) of the period end.
   */
  function summarise(result, periodStartSerial, periodEndSerial) {
    var emps = result.employees;
    var employed = emps.filter(function (e) {
      var engagedOk = e.engaged === null || e.engaged <= periodEndSerial;
      var notLeft = e.term === null || e.term > periodEndSerial;
      return engagedOk && notLeft;
    });
    var byLevel = { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0, unspecified: 0 };
    employed.forEach(function (e) { if (e.levelNo) byLevel[e.levelNo]++; else byLevel.unspecified++; });
    var fixIssues = result.issues.filter(function (i) { return i.field !== 'Left out'; });
    var issueCounts = {};
    fixIssues.forEach(function (i) { issueCounts[i.field] = (issueCounts[i.field] || 0) + 1; });
    var rowsWithIssues = {};
    fixIssues.forEach(function (i) { rowsWithIssues[i.src] = true; });
    return {
      employees: emps.length,
      employed: employed.length,
      terminatedBeforeEnd: emps.length - employed.length,
      byLevel: byLevel,
      issueCounts: issueCounts,
      issueCount: fixIssues.length,
      leftOut: (result.excluded || []).length,
      rowsWithIssues: Object.keys(rowsWithIssues).length,
      tooMany: emps.length > (LAST_ROW - FIRST_ROW + 1)
    };
  }

  // ---------------------------------------------------------------- template writing
  function sheetPath(fs, sheetName) {
    var wbXml = fs.read('xl/workbook.xml');
    var relsXml = fs.read('xl/_rels/workbook.xml.rels');
    var re = new RegExp('<sheet[^>]*name="' + sheetName.replace(/[&]/g, '&amp;') + '"[^>]*r:id="([^"]+)"');
    var m = wbXml.match(re);
    if (!m) throw new Error('Template has no sheet "' + sheetName + '"');
    var rel = new RegExp('<Relationship[^>]*Id="' + m[1] + '"[^>]*Target="([^"]+)"|<Relationship[^>]*Target="([^"]+)"[^>]*Id="' + m[1] + '"');
    var rm = relsXml.match(rel);
    var target = rm[1] || rm[2];
    target = target.replace(/^\//, '');
    return target.indexOf('xl/') === 0 ? target : 'xl/' + target;
  }

  /**
   * Set cell values in a sheet XML string. updates: {ref: value}; value null = blank cell.
   * Keeps each cell's existing style attribute. Rebuilds only the rows it touches.
   */
  function setCellsInSheet(xml, updates) {
    var byRow = {};
    Object.keys(updates).forEach(function (ref) {
      var mm = ref.match(/^([A-Z]+)(\d+)$/);
      (byRow[mm[2]] = byRow[mm[2]] || []).push({ col: mm[1], ref: ref, value: updates[ref] });
    });
    var touched = {};
    // one pass over the sheet: only the rows that carry updates are rebuilt
    var rowRe = /<row r="(\d+)"([^>]*?)(\/>|>([\s\S]*?)<\/row>)/g;
    var out = xml.replace(rowRe, function (whole, rowNum, attrs, tail, inner) {
      if (!byRow[rowNum]) return whole;
      touched[rowNum] = true;
      inner = inner || '';
      var cellRe = /<c r="([A-Z]+)\d+"([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
      var cells = {};
      var cm;
      while ((cm = cellRe.exec(inner)) !== null) cells[cm[1]] = { attrs: cm[2], raw: cm[0] };
      byRow[rowNum].forEach(function (u) {
        var old = cells[u.col];
        var sm = old ? old.attrs.match(/\ss="(\d+)"/) : null;
        var sAttr = sm ? ' s="' + sm[1] + '"' : '';
        var ref = u.col + rowNum;
        var v = u.value;
        var xmlCell;
        if (v === null || v === undefined || v === '') xmlCell = '<c r="' + ref + '"' + sAttr + '/>';
        else if (typeof v === 'number') xmlCell = '<c r="' + ref + '"' + sAttr + '><v>' + (Math.round(v * 1e6) / 1e6) + '</v></c>';
        else xmlCell = '<c r="' + ref + '"' + sAttr + ' t="inlineStr"><is><t xml:space="preserve">' + esc(v) + '</t></is></c>';
        cells[u.col] = { attrs: '', raw: xmlCell };
      });
      var ordered = Object.keys(cells).sort(function (a, b) { return colToNum(a) - colToNum(b); })
                           .map(function (k) { return cells[k].raw; }).join('');
      return '<row r="' + rowNum + '"' + attrs + '>' + ordered + '</row>';
    });
    Object.keys(byRow).forEach(function (rn) {
      if (!touched[rn]) throw new Error('Template row ' + rn + ' missing');
    });
    return out;
  }

  /** Build the list of cell updates for the data sheet and the settings sheet. */
  function buildUpdates(employees, settings) {
    var data = {}, sett = {};
    if (employees.length > (LAST_ROW - FIRST_ROW + 1)) throw new Error('More than ' + (LAST_ROW - FIRST_ROW + 1) + ' employees. Split the list or extend the template.');
    // how many times each code appears (calculated once here, so the workbook avoids a slow formula)
    var codeCount = {};
    employees.forEach(function (e) { if (e.code) codeCount[e.code] = (codeCount[e.code] || 0) + 1; });
    for (var rr = FIRST_ROW + employees.length; rr <= LAST_ROW; rr++) data['AZ' + rr] = null;
    employees.forEach(function (e, idx) {
      var r = FIRST_ROW + idx;
      data['AZ' + r] = e.code ? codeCount[e.code] : null;
      data['A' + r] = e.code;
      data['B' + r] = e.surname;
      data['C' + r] = e.names;
      data['D' + r] = e.gender;
      data['E' + r] = e.race;
      data['F' + r] = e.engaged;
      data['G' + r] = e.term;
      data['H' + r] = e.prevPos;
      data['I' + r] = e.curPos;
      data['J' + r] = e.level;
      data['K' + r] = e.perm;
      data['L' + r] = e.nationality;
      data['M' + r] = e.disabled;
      var cols = ['N','O','P','Q','R','S','T','U','V','W','X','Y'];
      for (var i = 0; i < 12; i++) data[cols[i] + r] = e.fixed[i];
      var vcols = ['Z','AA','AB','AC','AD','AE','AF'];
      for (var j = 0; j < 7; j++) data[vcols[j] + r] = e.variable[j];
      data['AH' + r] = e.levelStart;
    });
    sett.C3 = settings.company;
    sett.C4 = settings.periodStart;
    sett.C5 = settings.periodEnd;
    sett.C7 = settings.province;
    sett.C8 = settings.sector;
    sett.C10 = settings.semiMale === '' || settings.semiMale == null ? null : settings.semiMale;
    sett.C11 = settings.semiFemale === '' || settings.semiFemale == null ? null : settings.semiFemale;
    sett.C12 = settings.unskMale === '' || settings.unskMale == null ? null : settings.unskMale;
    sett.C13 = settings.unskFemale === '' || settings.unskFemale == null ? null : settings.unskFemale;
    return { data: data, settings: sett };
  }

  /** Add an "Import Log" sheet listing every issue, at the end of the workbook. */
  function addImportLog(fs, summary, issues, meta) {
    var wbPath = 'xl/workbook.xml';
    var wbXml = fs.read(wbPath);
    var relsPath = 'xl/_rels/workbook.xml.rels';
    var relsXml = fs.read(relsPath);
    var ctPath = '[Content_Types].xml';
    var ctXml = fs.read(ctPath);

    var existing = fs.names().filter(function (f) { return /^xl\/worksheets\/sheet\d+\.xml$/.test(f); });
    var maxN = 0;
    existing.forEach(function (f) { var n = +f.match(/sheet(\d+)\.xml$/)[1]; if (n > maxN) maxN = n; });
    var newFile = 'xl/worksheets/sheet' + (maxN + 1) + '.xml';
    var ids = (wbXml.match(/sheetId="(\d+)"/g) || []).map(function (s) { return +s.match(/\d+/)[0]; });
    var newSheetId = Math.max.apply(null, ids) + 1;
    var relIds = (relsXml.match(/Id="rId(\d+)"/g) || []).map(function (s) { return +s.match(/\d+/)[0]; });
    var newRid = 'rId' + (Math.max.apply(null, relIds) + 1);

    // workbook.xml: append the sheet at the end (no index shift for defined names)
    wbXml = wbXml.replace('</sheets>', function () { return '<sheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" name="Import Log" sheetId="' + newSheetId + '" r:id="' + newRid + '"/></sheets>'; });
    fs.write(wbPath, wbXml);
    relsXml = relsXml.replace('</Relationships>', '<Relationship Id="' + newRid + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + (maxN + 1) + '.xml"/></Relationships>');
    fs.write(relsPath, relsXml);
    ctXml = ctXml.replace('</Types>', '<Override PartName="/' + newFile + '" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>');
    fs.write(ctPath, ctXml);

    // rows of text
    var lines = [];
    lines.push(['IMPORT LOG: data problems found when the employee list was loaded']);
    lines.push(['Company: ' + meta.company + '   Period: ' + meta.periodText]);
    lines.push(['Employees loaded: ' + summary.employees + '   Employed on period end: ' + summary.employed + '   Issues: ' + summary.issueCount + ' on ' + summary.rowsWithIssues + ' rows']);
    lines.push(['Fix the items below in the source list (or in the workbook), then the Checks sheet will clear.']);
    (meta.notes || []).forEach(function (n) { lines.push(['Note: ' + n]); });
    lines.push([]);
    lines.push(['Source row', 'Empl code', 'Name', 'Field', 'Problem', 'Action needed']);
    issues.forEach(function (i) { lines.push([i.src, i.code, i.name, i.field, i.problem, i.action]); });

    var bold = boldStyleId(fs);
    var sAttr = function (rn) { return bold && (rn === 1 || rn === 6) ? ' s="' + bold + '"' : ''; };
    var rowsXml = lines.map(function (cells, idx) {
      var rn = idx + 1;
      var cx = cells.map(function (v, ci) {
        var ref = numToCol(ci + 1) + rn;
        if (v === '' || v === null || v === undefined) return '';
        if (typeof v === 'number') return '<c r="' + ref + '"' + sAttr(rn) + '><v>' + v + '</v></c>';
        return '<c r="' + ref + '"' + sAttr(rn) + ' t="inlineStr"><is><t xml:space="preserve">' + esc(v) + '</t></is></c>';
      }).join('');
      return cx ? '<row r="' + rn + '">' + cx + '</row>' : '';
    }).join('');
    var sheetXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheetViews><sheetView workbookViewId="0"><pane ySplit="6" topLeftCell="A7" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>' +
      '<sheetFormatPr defaultRowHeight="15"/>' +
      '<cols><col min="1" max="1" width="12" customWidth="1"/><col min="2" max="2" width="16" customWidth="1"/><col min="3" max="3" width="30" customWidth="1"/><col min="4" max="4" width="20" customWidth="1"/><col min="5" max="5" width="60" customWidth="1"/><col min="6" max="6" width="48" customWidth="1"/></cols>' +
      '<sheetData>' + rowsXml + '</sheetData>' +
      '<autoFilter ref="A6:F' + lines.length + '"/></worksheet>';
    fs.write(newFile, sheetXml);
    // Excel needs this hidden name to show the filter buttons on the Import Log
    var sheetIndex = (wbXml.match(/<sheet\b/g) || []).length - 1;
    var filterName = '<definedName name="_xlnm._FilterDatabase" localSheetId="' + sheetIndex + '" hidden="1">\'Import Log\'!$A$6:$F$' + lines.length + '</definedName>';
    if (wbXml.indexOf('<definedNames>') !== -1) wbXml = wbXml.replace('</definedNames>', function () { return filterName + '</definedNames>'; });
    else wbXml = wbXml.replace('</sheets>', function () { return '</sheets><definedNames>' + filterName + '</definedNames>'; });
    fs.write(wbPath, wbXml);
  }

  /** Build the Positions sheet: every current position in the list, with live counts and a filter. */
  function addPositions(fs, employees, meta) {
    if (!employees.length) return;
    var groups = {};
    employees.forEach(function (e) {
      var disp = cleanText(e.curPos);
      var key = disp === '' ? '' : disp.toLowerCase();
      if (!groups[key]) groups[key] = { counts: {}, total: 0 };
      var g = groups[key];
      g.total++;
      g.counts[disp] = (g.counts[disp] || 0) + 1;
    });
    var keys = Object.keys(groups).filter(function (k) { return k !== ''; });
    keys.sort(function (a, b) { return a < b ? -1 : a > b ? 1 : 0; });
    var rowsData = keys.map(function (k) {
      var best = Object.keys(groups[k].counts).sort(function (a, b) { return groups[k].counts[b] - groups[k].counts[a]; })[0];
      return { text: best };
    });
    var hasBlank = groups[''] !== undefined;

    var DATA = "'EEA2 & EEA4 INFO'!";
    var I = DATA + '$I$10:$I$' + LAST_ROW, AQ = DATA + '$AQ$10:$AQ$' + LAST_ROW, AP = DATA + '$AP$10:$AP$' + LAST_ROW;
    var first = 6, last = first + rowsData.length - 1 + (hasBlank ? 1 : 0);
    var lines = [];
    lines.push([{ v: 'CURRENT POSITIONS: every position in the list. Filter the headings to choose one.' }]);
    lines.push([{ v: 'Company: ' + meta.company + '    Period: ' + meta.periodText }]);
    lines.push([]);
    lines.push([{ v: 'All positions' },
                { f: 'SUM(B' + first + ':B' + last + ')' }, { f: 'SUM(C' + first + ':C' + last + ')' },
                { f: 'SUM(D' + first + ':D' + last + ')' }, { f: 'SUM(E' + first + ':E' + last + ')' }]);
    lines.push([{ v: 'Current position' }, { v: 'Employees on list' }, { v: 'Employed at period end' },
                { v: 'Permanent (employed)' }, { v: 'Temporary (employed)' }, { v: 'helper' }]);
    rowsData.forEach(function (rd, i) {
      var r = first + i;
      lines.push([{ v: rd.text },
        { f: 'COUNTIFS(' + I + ',"="&$F' + r + ')' },
        { f: 'COUNTIFS(' + I + ',"="&$F' + r + ',' + AQ + ',1)' },
        { f: 'COUNTIFS(' + I + ',"="&$F' + r + ',' + AQ + ',1,' + AP + ',0)' },
        { f: 'COUNTIFS(' + I + ',"="&$F' + r + ',' + AQ + ',1,' + AP + ',1)' },
        { f: 'SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(A' + r + ',"~","~~"),"*","~*"),"?","~?")' }]);
    });
    if (hasBlank) {
      var rb = first + rowsData.length;
      var prevSum = function (col) { return 'SUM(' + col + first + ':' + col + (rb - 1) + ')'; };
      lines.push([{ v: '(no current position entered)' },
        { f: 'COUNT(' + AQ + ')-' + prevSum('B') },
        { f: 'SUM(' + AQ + ')-' + prevSum('C') },
        { f: 'COUNTIFS(' + AQ + ',1,' + AP + ',0)-' + prevSum('D') },
        { f: 'COUNTIFS(' + AQ + ',1,' + AP + ',1)-' + prevSum('E') },
        { v: '' }]);
    }
    var colL = ['A', 'B', 'C', 'D', 'E', 'F'];
    var bold = boldStyleId(fs);
    var sAttr = function (rn) { return bold && (rn === 1 || rn === 4 || rn === 5) ? ' s="' + bold + '"' : ''; };
    var xmlRows = lines.map(function (cells, idx) {
      var rn = idx + 1;
      var cx = cells.map(function (c, ci) {
        var ref = colL[ci] + rn;
        if (c.f !== undefined) return '<c r="' + ref + '"' + sAttr(rn) + '><f>' + escXml(c.f) + '</f><v></v></c>';
        if (c.v === undefined || c.v === '') return '';
        if (typeof c.v === 'number') return '<c r="' + ref + '"' + sAttr(rn) + '><v>' + c.v + '</v></c>';
        return '<c r="' + ref + '"' + sAttr(rn) + ' t="inlineStr"><is><t xml:space="preserve">' + escXml(c.v) + '</t></is></c>';
      }).join('');
      return cx ? '<row r="' + rn + '">' + cx + '</row>' : '';
    }).join('');
    var headerRow = 5;
    var lastRow = lines.length;
    var sheetXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheetViews><sheetView workbookViewId="0"><pane ySplit="' + headerRow + '" topLeftCell="A' + (headerRow + 1) + '" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>' +
      '<sheetFormatPr defaultRowHeight="15"/>' +
      '<cols><col min="1" max="1" width="48" customWidth="1"/><col min="2" max="5" width="22" customWidth="1"/><col min="6" max="6" width="10" hidden="1" customWidth="1"/></cols>' +
      '<sheetData>' + xmlRows + '</sheetData>' +
      '<autoFilter ref="A' + headerRow + ':E' + lastRow + '"/></worksheet>';
    addSheet(fs, 'Positions', sheetXml, headerRow, lastRow);
  }

  function boldStyleId(fs) {
    var m = fs.read(sheetPath(fs, 'EEA12 Analysis')).match(/<c r="A5"[^>]*?\ss="(\d+)"/);
    return m ? m[1] : null;
  }

  function escXml(t) {
    return String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  /** Append a worksheet part to the workbook, with its filter name. */
  function addSheet(fs, name, sheetXml, headerRow, lastRow) {
    var wbPath = 'xl/workbook.xml', relsPath = 'xl/_rels/workbook.xml.rels', ctPath = '[Content_Types].xml';
    var wbXml = fs.read(wbPath), relsXml = fs.read(relsPath), ctXml = fs.read(ctPath);
    var existing = fs.names().filter(function (f) { return /^xl\/worksheets\/sheet\d+\.xml$/.test(f); });
    var maxN = 0;
    existing.forEach(function (f) { var n = +f.match(/sheet(\d+)\.xml$/)[1]; if (n > maxN) maxN = n; });
    var newFile = 'xl/worksheets/sheet' + (maxN + 1) + '.xml';
    var ids = (wbXml.match(/sheetId="(\d+)"/g) || []).map(function (x) { return +x.match(/\d+/)[0]; });
    var newSheetId = Math.max.apply(null, ids) + 1;
    var relIds = (relsXml.match(/Id="rId(\d+)"/g) || []).map(function (x) { return +x.match(/\d+/)[0]; });
    var newRid = 'rId' + (Math.max.apply(null, relIds) + 1);
    wbXml = wbXml.replace('</sheets>', function () { return '<sheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" name="' + name + '" sheetId="' + newSheetId + '" r:id="' + newRid + '"/></sheets>'; });
    var idx = (wbXml.match(/<sheet\b/g) || []).length - 1;
    var fname = '<definedName name="_xlnm._FilterDatabase" localSheetId="' + idx + '" hidden="1">\'' + name + '\'!$A$' + headerRow + ':$E$' + lastRow + '</definedName>';
    if (wbXml.indexOf('<definedNames>') !== -1) wbXml = wbXml.replace('</definedNames>', function () { return fname + '</definedNames>'; });
    else wbXml = wbXml.replace('</sheets>', function () { return '</sheets><definedNames>' + fname + '</definedNames>'; });
    fs.write(wbPath, wbXml);
    relsXml = relsXml.replace('</Relationships>', function () { return '<Relationship Id="' + newRid + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + (maxN + 1) + '.xml"/></Relationships>'; });
    fs.write(relsPath, relsXml);
    ctXml = ctXml.replace('</Types>', function () { return '<Override PartName="/' + newFile + '" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'; });
    fs.write(ctPath, ctXml);
    fs.write(newFile, sheetXml);
  }

  /**
   * Build the complete workbook.
   *  templateBytes: Uint8Array of the template .xlsx
   *  result: normalised employees + issues (after addDuplicateIssues)
   *  settings: {company, periodStart (serial), periodEnd (serial), periodText, province, sector,
   *             semiMale, semiFemale, unskMale, unskFemale}
   * Returns Promise<Uint8Array>.
   */
  function buildWorkbook(templateBytes, result, settings, summary) {
    var zip, fs;
    return JSZip.loadAsync(templateBytes).then(function (z) {
      zip = z;
      var names = Object.keys(zip.files).filter(function (n) { return !zip.files[n].dir; });
      var jobs = names.filter(function (n) { return /\.(xml|rels)$/.test(n); })
                      .map(function (n) { return zip.file(n).async('string').then(function (t) { return [n, t]; }); });
      return Promise.all(jobs);
    }).then(function (pairs) {
      var store = {};
      pairs.forEach(function (p) { store[p[0]] = p[1]; });
      fs = {
        read: function (n) { if (!(n in store)) throw new Error('Template part missing: ' + n); return store[n]; },
        write: function (n, t) { store[n] = t; },
        names: function () { return Object.keys(store); },
        store: store
      };
      var upd = buildUpdates(result.employees, settings);
      var dataPath = sheetPath(fs, DATA_SHEET);
      var setPath = sheetPath(fs, 'Settings');
      fs.write(dataPath, setCellsInSheet(fs.read(dataPath), upd.data));
      fs.write(setPath, setCellsInSheet(fs.read(setPath), upd.settings));
      addImportLog(fs, summary, result.issues.concat(result.excluded || []), { company: settings.company, periodText: settings.periodText, notes: result.notes || [] });
      addPositions(fs, result.employees, { company: settings.company, periodText: settings.periodText });
      // Ask Excel to recalculate everything when the file opens
      var wbXml = fs.read('xl/workbook.xml');
      if (wbXml.indexOf('fullCalcOnLoad') === -1) {
        wbXml = wbXml.replace(/<calcPr([^>]*?)\/>/, '<calcPr$1 fullCalcOnLoad="1"/>');
        fs.write('xl/workbook.xml', wbXml);
      }
      // Calculate every formula and store the results, so the figures show before editing is enabled
      var wbNow = fs.read('xl/workbook.xml');
      var sheetList = [];
      (wbNow.match(/<sheet\b[^>]*\/>/g) || []).forEach(function (tag) {
        var nm = tag.match(/name="([^"]+)"/)[1].replace(/&amp;/g, '&');
        sheetList.push({ name: nm, path: sheetPath(fs, nm) });
      });
      var calcSheets = sheetList.map(function (x) { return { name: x.name, xml: fs.read(x.path) }; });
      var results = EECalc.calculate(calcSheets);
      sheetList.forEach(function (x) {
        if (results[x.name] && results[x.name].length) fs.write(x.path, EECalc.store(fs.read(x.path), results[x.name]));
      });
      Object.keys(fs.store).forEach(function (n) { zip.file(n, fs.store[n], { createFolders: false }); });
      return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
    });
  }

  /** Read the province and sector drop-down lists from the template. */
  function readTemplateLists(templateBytes) {
    var wb = XLSX.read(templateBytes, { type: 'array' });
    function col(sheet, c, r1, r2) {
      var out = [];
      for (var r = r1; r <= r2; r++) {
        var cell = wb.Sheets[sheet][c + r];
        if (cell && cell.v) out.push(String(cell.v));
      }
      return out;
    }
    return { provinces: col('EAP Targets', 'B', 4, 13), sectors: col('Sector Targets', 'B', 4, 21) };
  }

  /** Period text like "1 September 2025 to 31 August 2026" from the title of a skills report. */
  function detectPeriod(rows) {
    var out = { company: '', periodText: '' };
    for (var r = 0; r < Math.min(rows.length, 15); r++) {
      var t = rows[r][0];
      if (typeof t !== 'string') continue;
      var mc = t.match(/SKILLS REPORT FOR\s+(.+?)\s*(\n|$)/i);
      if (mc && !out.company) out.company = mc[1].trim();
      var mp = t.match(/(\d{1,2}\s+[A-Za-z]+\s+\d{4})\s*-\s*(\d{1,2}\s+[A-Za-z]+\s+\d{4})/);
      if (mp && !out.periodText) out.periodText = mp[1] + ' to ' + mp[2];
      var mi = String(rows[r][1] || '').match(/from\s+(\d{1,2}\s+[A-Za-z]+\s+\d{4})\s+to\s+(\d{1,2}\s+[A-Za-z]+\s+\d{4})/i);
      if (mi && !out.periodText) out.periodText = mi[1] + ' to ' + mi[2];
    }
    return out;
  }

  /** Parse "1 September 2025" style text to an Excel serial, or null. */
  function parseLongDate(text) {
    var m = String(text || '').match(/(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})/);
    if (!m) return null;
    var months = ['january','february','march','april','may','june','july','august','september','october','november','december'];
    var mi = months.indexOf(m[2].toLowerCase());
    if (mi < 0) return null;
    return serialFromYMD(+m[3], mi + 1, +m[1]);
  }

  return {
    FIRST_ROW: FIRST_ROW, LAST_ROW: LAST_ROW, LEVEL_TEXT: LEVEL_TEXT,
    openWorkbook: openWorkbook, readRows: readRows, detectHeader: detectHeader, normalise: normalise,
    addDuplicateIssues: addDuplicateIssues, summarise: summarise,
    buildWorkbook: buildWorkbook, readTemplateLists: readTemplateLists,
    detectPeriod: detectPeriod, parseLongDate: parseLongDate, readPermMap: readPermMap, applyPeriod: applyPeriod,
    parseDate: parseDate, parseNumber: parseNumber, detectDateOrder: detectDateOrder,
    setCellsInSheet: setCellsInSheet, serialFromYMD: serialFromYMD
  };
});
