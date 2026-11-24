/*
 * Browser logic for the Employment Equity Workbook Builder.
 * File contents are only ever written with textContent (never innerHTML), so names
 * from an uploaded file cannot inject markup into the page.
 */
(function () {
  'use strict';

  var MAX_ISSUES_SHOWN = 300;
  var FIELD_LABELS = {
    code: 'Employee code', surname: 'Surname', names: 'Full names', gender: 'Gender', race: 'Race',
    engaged: 'Date engaged', term: 'Termination date', prev_pos: 'Previous position',
    cur_pos: 'Current position', level: 'Occupational level', level_no: 'Level number',
    level_start: 'Level at start of period', perm: 'Permanent / temporary', nationality: 'Nationality',
    fn_flag: 'Foreign national flag (FN)', disabled: 'Disabled', total_rem: 'Total remuneration'
  };
  var REQUIRED_LABEL = {
    code: 'Employee code', gender: 'Gender', race: 'Race', engaged: 'Date engaged', level: 'Occupational level'
  };

  var state = { template: null, lists: null, parsed: null };
  function $(id) { return document.getElementById(id); }

  function el(tag, cls, text) {
    var n = document.createElement(tag);
    if (cls) n.className = cls;
    if (text !== undefined && text !== null) n.textContent = String(text);
    return n;
  }
  function clear(node) { while (node.firstChild) node.removeChild(node.firstChild); }

  function b64ToBytes(b64) {
    var bin = atob(b64), out = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return out;
  }

  function showError(msg) {
    var box = $('fileError');
    box.textContent = msg;
    box.hidden = !msg;
  }

  function fmtDate(iso) {
    var p = iso.split('-');
    var months = ['January','February','March','April','May','June','July','August','September','October','November','December'];
    return (+p[2]) + ' ' + months[+p[1] - 1] + ' ' + p[0];
  }

  function periodSerials() {
    var s = $('periodStart').value, e = $('periodEnd').value;
    if (!s || !e) return null;
    var a = s.split('-').map(Number), b = e.split('-').map(Number);
    var ps = EECore.serialFromYMD(a[0], a[1], a[2]);
    var pe = EECore.serialFromYMD(b[0], b[1], b[2]);
    if (!(pe > ps)) return null;
    return { start: ps, end: pe, text: fmtDate(s) + ' to ' + fmtDate(e), startIso: s, endIso: e };
  }

  // ------------------------------------------------------------- init
  function init() {
    state.template = b64ToBytes(window.EE_TEMPLATE_B64);
    state.lists = EECore.readTemplateLists(state.template);
    fillSelect($('province'), state.lists.provinces, 'Choose the EAP');
    fillSelect($('sector'), state.lists.sectors, 'Choose the sector');
    $('province').value = '';
    $('sector').value = '';

    var fileInput = $('file'), drop = $('drop');
    fileInput.addEventListener('change', function () { if (fileInput.files[0]) loadFile(fileInput.files[0]); });
    ['dragenter', 'dragover'].forEach(function (ev) {
      drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.add('over'); });
    });
    ['dragleave', 'drop'].forEach(function (ev) {
      drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.remove('over'); });
    });
    drop.addEventListener('drop', function (e) {
      var f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
      if (f) loadFile(f);
    });

    $('dateFormat').addEventListener('change', function () {
      if (!state.parsed) return;
      normaliseNow(this.value);
      renderMapping();
      refresh();
    });
    ['company', 'periodStart', 'periodEnd', 'province', 'sector', 'semiMale', 'semiFemale', 'unskMale', 'unskFemale']
      .forEach(function (id) {
        $(id).addEventListener('input', refresh);
        $(id).addEventListener('change', refresh);
      });
    $('build').addEventListener('click', buildWorkbook);
    refresh();
  }

  function fillSelect(sel, items, placeholder) {
    clear(sel);
    sel.appendChild(el('option', null, placeholder)).value = '';
    items.forEach(function (t) {
      var o = el('option', null, t); o.value = t; sel.appendChild(o);
    });
  }

  // ------------------------------------------------------------- loading
  function loadFile(file) {
    showError('');
    $('dropText').textContent = 'Reading ' + file.name + ' ...';
    var reader = new FileReader();
    reader.onerror = function () { showError('The file could not be read.'); };
    reader.onload = function () {
      try {
        processBuffer(reader.result, file.name);
        $('dropText').textContent = 'Loaded ' + file.name + '. Drop another file to replace it.';
      } catch (err) {
        state.parsed = null;
        $('dropText').textContent = 'Drop the file here, or click to choose it';
        showError(err && err.message ? err.message : String(err));
        $('step-map').hidden = true;
        $('step-check').hidden = true;
        refresh();
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function chooseSheet(buf) {
    // Try every sheet; use the one with the most employee rows (ties go to the first sheet).
    var names = EECore.openWorkbook(buf).SheetNames;
    var best = null, reasons = [];
    for (var i = 0; i < names.length; i++) {
      try {
        var rd = EECore.readRows(buf, names[i]);
        var det = EECore.detectHeader(rd.rows);
        if (det.fatal.length) {
          reasons.push('"' + names[i] + '": missing ' + det.fatal.map(function (k) { return REQUIRED_LABEL[k] || k; }).join(', '));
          continue;
        }
        var res = EECore.normalise(rd.rows, det);
        if (res.employees.length === 0) { reasons.push('"' + names[i] + '": header found but no employee rows under it'); continue; }
        if (!best || res.employees.length > best.count) best = { name: names[i], rd: rd, det: det, count: res.employees.length };
      } catch (e) {
        reasons.push('"' + names[i] + '": ' + e.message);
      }
    }
    if (!best) throw new Error('No sheet in this file has the employee columns this tool needs. ' + reasons.join(' | '));
    return best;
  }

  function processBuffer(buf, fileName) {
    var chosen = chooseSheet(buf);
    var rd = chosen.rd, det = chosen.det;
    var per = EECore.detectPeriod(rd.rows);
    var pm = EECore.readPermMap(buf, chosen.name);
    state.parsed = { fileName: fileName, sheet: chosen.name, rows: rd.rows, det: det, pm: pm, per: per, res: null };
    $('dateFormat').value = 'auto';
    normaliseNow($('dateFormat').value);
    var res = state.parsed.res;

    // Pre-fill company and period from the file title when the user has not typed anything yet
    if (per.company && !$('company').value) $('company').value = per.company.replace(/\b\w/g, function (c) { return c.toUpperCase(); });
    var m = (per.periodText || '').match(/(\d{1,2}\s+[A-Za-z]+\s+\d{4})\s+to\s+(\d{1,2}\s+[A-Za-z]+\s+\d{4})/);
    if (m) {
      var s = EECore.parseLongDate(m[1]), e = EECore.parseLongDate(m[2]);
      if (s && e) {
        $('periodStart').value = isoFromSerial(s);
        $('periodEnd').value = isoFromSerial(e);
      }
    }
    renderMapping();
    $('step-map').hidden = false;
    $('fileInfo').hidden = false;
    $('fileInfo').textContent = fileName + ': sheet "' + chosen.name + '", header on row ' + (det.headerRow + 1) +
      ', ' + res.employees.length + ' employee rows read.';
    refresh();
  }

  /** Read the employee rows again with the chosen date format. */
  function normaliseNow(fmt) {
    var p = state.parsed;
    p.res = EECore.addDuplicateIssues(EECore.normalise(p.rows, p.det, p.pm ? p.pm.map : null, { dateFormat: fmt }));
    var info = (p.res.notes || []).filter(function (n) { return /^Text dates read/.test(n); })[0] || '';
    $('dateInfo').textContent = info;
  }

  function isoFromSerial(serial) {
    var ms = Math.round((serial - 25569) * 86400000), d = new Date(ms);
    var pad = function (n) { return (n < 10 ? '0' : '') + n; };
    return d.getUTCFullYear() + '-' + pad(d.getUTCMonth() + 1) + '-' + pad(d.getUTCDate());
  }

  // ------------------------------------------------------------- mapping table
  function renderMapping() {
    var p = state.parsed, det = p.det, rows = p.rows;
    var header = rows[det.headerRow] || [];
    var table = $('mapTable');
    clear(table);
    var head = el('tr');
    ['Field in the workbook', 'Column found in your file', 'Status'].forEach(function (h) { head.appendChild(el('th', null, h)); });
    table.appendChild(head);

    Object.keys(FIELD_LABELS).forEach(function (f) {
      var tr = el('tr');
      var col = det.map[f];
      tr.appendChild(el('td', null, FIELD_LABELS[f]));
      tr.appendChild(el('td', null, col === undefined ? '-' : String(header[col]).replace(/\s+/g, ' ').trim()));
      var st = el('td');
      var optional = ['level_no', 'level_start', 'fn_flag', 'total_rem'].indexOf(f) !== -1;
      if (col !== undefined) st.appendChild(el('span', 'tag ok', 'found'));
      else if (det.fatal.indexOf(f) !== -1) st.appendChild(el('span', 'tag bad', 'required: missing'));
      else if (f === 'total_rem' && det.fixedCols.length) st.appendChild(el('span', 'tag muted', 'not needed'));
      else if (optional) st.appendChild(el('span', 'tag muted', 'optional'));
      else st.appendChild(el('span', 'tag warn', 'missing: entered as blank'));
      tr.appendChild(st);
      table.appendChild(tr);
    });
    var rem = el('tr');
    rem.appendChild(el('td', null, 'Remuneration (fixed and variable)'));
    rem.appendChild(el('td', null, det.fixedCols.length + ' fixed and ' + det.varCols.length + ' variable columns'));
    var remSt = el('td');
    remSt.appendChild(det.fixedCols.length ? el('span', 'tag ok', 'found') :
      el('span', 'tag warn', det.map.total_rem !== undefined ? 'total only' : 'missing'));
    rem.appendChild(remSt);
    table.appendChild(rem);

    var hint = $('mapHint');
    hint.textContent = det.fatal.length
      ? 'This file cannot be used until the required columns are present. See the status column.'
      : 'Columns are matched by their names, so the order of columns does not matter.';
  }

  // ------------------------------------------------------------- check
  function renderCheck() {
    var p = state.parsed, per = periodSerials();
    var card = $('step-check');
    if (!p) { card.hidden = true; return; }
    card.hidden = false;
    var res = per ? EECore.applyPeriod(p.res, per.start, per.end) : p.res;
    var sum = per ? EECore.summarise(res, per.start, per.end) : null;

    var cards = $('summaryCards');
    clear(cards);
    function stat(n, label, cls) {
      var d = el('div', 'stat ' + (cls || ''));
      d.appendChild(el('div', 'n', n));
      d.appendChild(el('div', 'l', label));
      cards.appendChild(d);
    }
    stat(res.employees.length, 'employees in the list');
    stat(sum ? sum.employed : '-', 'employed on period end', 'ok');
    stat(sum ? sum.terminatedBeforeEnd : '-', 'left during the period');
    stat(sum ? sum.byLevel.unspecified : '-', 'employed with no level', sum && sum.byLevel.unspecified ? 'warn' : 'ok');
    var fixes = res.issues.filter(function (i) { return i.field !== 'Left out'; });
    stat(fixes.length, 'items to fix (' + new Set(fixes.map(function (i) { return i.src; })).size + ' rows)',
         fixes.length ? 'bad' : 'ok');


    var notes = $('notes');
    clear(notes);
    (res.notes || []).forEach(function (n) { notes.appendChild(el('p', 'info', n)); });
    var leftOut = { contractors: 0, period: 0 };
    (p.res.excluded || []).forEach(function (x) {
      if (x.problem.indexOf('independent contractor') === 0) leftOut.contractors++;
    });
    if (per) (EECore.applyPeriod(p.res, per.start, per.end).excluded || []).forEach(function (x) {
      if (x.problem.indexOf('not employed during the period') === 0) leftOut.period++;
    });
    if (leftOut.contractors) notes.appendChild(el('p', 'info', leftOut.contractors + ' independent contractor row(s) left out (not employees under the EEA).'));
    if (leftOut.period) notes.appendChild(el('p', 'info', leftOut.period + ' row(s) left out because the person was not employed during the period (engaged after the period end, or left before it started). They are listed on the Import Log.'));
    var absent = p.det.missing.filter(function (k) {
      if (k === 'nationality' && p.det.map.fn_flag !== undefined) return false;   // FN column covers it
      return true;
    });
    if (absent.length && !p.det.fatal.length) {
      notes.appendChild(el('p', 'info', 'Not in this file (left blank in the workbook): ' +
        absent.map(function (k) { return FIELD_LABELS[k] || k; }).join(', ') + '.'));
    }
    if (sum && sum.tooMany) notes.appendChild(el('p', 'alert', 'The workbook holds up to ' + (EECore.LAST_ROW - EECore.FIRST_ROW + 1) + ' employees. This list has ' + res.employees.length + ' in the period. Split the list first.'));
    if (!per) notes.appendChild(el('p', 'alert', 'The period end must be after the period start.'));

    // issues by field
    var byField = {};
    res.issues.forEach(function (i) { byField[i.field] = (byField[i.field] || 0) + 1; });
    var fieldKeys = Object.keys(byField).sort(function (a, b) { return byField[b] - byField[a]; });
    $('issueHint').textContent = res.issues.length
      ? fieldKeys.map(function (k) { return k + ': ' + byField[k]; }).join('   |   ') +
        '. The workbook is still created with these items flagged on the Import Log tab.'
      : 'No problems found. Every employee has the details the workbook needs.';

    var table = $('issueTable');
    clear(table);
    if (res.issues.length) {
      var head = el('tr');
      ['Source row', 'Empl code', 'Name', 'Field', 'Problem', 'Action needed'].forEach(function (h) { head.appendChild(el('th', null, h)); });
      table.appendChild(head);
      res.issues.slice(0, MAX_ISSUES_SHOWN).forEach(function (i) {
        var tr = el('tr');
        [i.src, i.code, i.name, i.field, i.problem, i.action].forEach(function (v) { tr.appendChild(el('td', null, v)); });
        table.appendChild(tr);
      });
    }
    $('issueMore').textContent = res.issues.length > MAX_ISSUES_SHOWN
      ? 'Showing the first ' + MAX_ISSUES_SHOWN + ' of ' + res.issues.length + '. The full list is on the Import Log tab of the workbook.'
      : (res.issues.length ? 'The full list is on the Import Log tab of the workbook.' : '');
  }

  // ------------------------------------------------------------- refresh + build
  function settingsValid() {
    var p = state.parsed;
    var per = periodSerials();
    var problems = [];
    if (!p) problems.push('upload the employee list');
    else if (p.det.fatal.length) problems.push('fix the required columns in the file');
    if (!$('company').value.trim()) problems.push('enter the company name');
    if (!per) problems.push('set the period (end after start)');
    if (!$('province').value) problems.push('choose the EAP');
    if (!$('sector').value) problems.push('choose the sector');
    if (p && per && EECore.applyPeriod(p.res, per.start, per.end).employees.length > (EECore.LAST_ROW - EECore.FIRST_ROW + 1)) problems.push('the list is too long for one workbook');
    ['semiMale', 'semiFemale', 'unskMale', 'unskFemale'].forEach(function (id) {
      var v = $(id).value;
      if (v !== '' && (isNaN(+v) || +v < 0 || +v > 100)) problems.push('goals must be between 0 and 100');
    });
    return problems;
  }

  function refresh() {
    renderCheck();
    var problems = settingsValid();
    $('build').disabled = problems.length > 0;
    var msg = $('buildMsg');
    msg.className = 'msg';
    msg.textContent = problems.length ? 'To create the workbook: ' + problems.join('; ') + '.' : 'Ready. Click Create workbook.';
  }

  function pctToFraction(id) {
    var v = $(id).value;
    if (v === '') return '';
    return Math.round(+v * 10) / 1000;   // percent to fraction, 0.1% resolution
  }

  function buildWorkbook() {
    var problems = settingsValid();
    if (problems.length) return refresh();
    var p = state.parsed, per = periodSerials();
    var eff = EECore.applyPeriod(p.res, per.start, per.end);
    var sum = EECore.summarise(eff, per.start, per.end);
    var settings = {
      company: $('company').value.trim(),
      periodStart: per.start, periodEnd: per.end, periodText: per.text,
      province: $('province').value, sector: $('sector').value,
      semiMale: pctToFraction('semiMale'), semiFemale: pctToFraction('semiFemale'),
      unskMale: pctToFraction('unskMale'), unskFemale: pctToFraction('unskFemale')
    };
    var msg = $('buildMsg');
    msg.className = 'msg';
    msg.textContent = 'Building the workbook ...';
    $('build').disabled = true;

    EECore.buildWorkbook(state.template, eff, settings, sum).then(function (bytes) {
      var blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      var name = 'EE Workbook - ' + settings.company.replace(/[^A-Za-z0-9 _-]+/g, '').trim() + ' - ' +
                 per.startIso.slice(0, 4) + '-' + per.endIso.slice(0, 4) + '.xlsx';
      var url = URL.createObjectURL(blob);
      var a = el('a');
      a.href = url; a.download = name;
      document.body.appendChild(a); a.click();
      setTimeout(function () { URL.revokeObjectURL(url); a.remove(); }, 2000);
      msg.className = 'msg ok';
      msg.textContent = 'Downloaded ' + name + '. Open it in Excel; the figures calculate when it opens.';
      $('build').disabled = false;
    }).catch(function (err) {
      msg.className = 'msg err';
      msg.textContent = 'The workbook could not be created: ' + (err && err.message ? err.message : err);
      $('build').disabled = false;
    });
  }

  document.addEventListener('DOMContentLoaded', init);
})();
