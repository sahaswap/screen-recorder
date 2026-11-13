/*
  ECM Transactions Export
  For each alert ID (ECM case ID):
    1. reads Customer Name, Customer Internal ID, current and permanent address, and Velocity Alert ID from the Requirements page;
    2. runs "Export Transactions" on the Statistics page, saving the file to Downloads as
       <ECM ID>_<Velocity Alert ID>_<original file name>;
    3. reads the Alert Analyst Decision and the Alert Narrative from the Requirements page;
  then saves ECM_Alert_Info_<date>.csv (one row per alert) for the Excel tool, and
  ECM_Narratives_<date>.xlsx (one row per alert, the narrative in full).
  Run on any ECM page while logged in: paste into the Console (Ctrl+Shift+J) or use the bookmarklet.
  Uses your existing Edge session - no passwords are stored or sent anywhere else.
*/
(function () {
  if (document.getElementById('__ecmExport')) return;

  var VERSION = '2026-11-13 (with Narratives)';   /* shown in the panel title */
  var BASE = location.origin + '/ECM/Case/Statistics.aspx?caseid=';
  var REQ_BASE = location.origin + '/ECM/Case/Requirements.aspx?caseid=';
  /* Focal Entity info tab; the second address is tried if the first does not answer. */
  var FOCAL_BASES = [location.origin + '/ECM/Case/CustomerInfo?caseid=', location.origin + '/ECM/Case/FocalEntity.aspx?caseid='];
  /* Off-screen but rendered: page text is only readable from a frame that is not hidden. */
  var FRAME_CSS = 'position:fixed;left:-10000px;top:0;width:1200px;height:800px;border:0;opacity:0;pointer-events:none';
  var COLUMNS = ['ECM ID', 'Velocity Alert ID', 'Customer Name', 'Customer Internal ID', 'Address',
    'Street', 'City', 'State', 'Zip Code', 'Country', 'Permanent Address',
    'Program ID', 'Program Name', 'Program Description', 'Transaction File', 'Result'];
  var results = [];
  var BUTTON_ID = 'btnExportStatsTransactions';
  var CLICK_WAIT_MS = 10000; /* after clicking Export: time allowed for the download to start */
  var PAUSE_MS = 300;      /* pause between alerts, to go easy on the server */
  var POLL_MS = 200;       /* how often a loading page is checked */
  var PAGE_TIMEOUT_MS = 60000;
  var PARALLEL = 2;        /* default alerts worked on at the same time; the panel's picker overrides it */
  try { PARALLEL = +localStorage.getItem('__ecmParallel') || PARALLEL; } catch (e) { /* storage blocked */ }
  var stopped = false;
  var running = false;

  function sleep(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }

  function unique(list) { return list.filter(function (v, i) { return list.indexOf(v) === i; }); }

  function parseIds(text) { return unique(text.match(/\d{3,}/g) || []); }

  /* Alert IDs from the "My Case Queue" table on the Dashboard. */
  function queueIds() {
    var links = document.querySelectorAll('table[id^="gvMyCaseQueue"] a[href*="caseid="]');
    return unique(Array.prototype.map.call(links, function (a) {
      var m = /caseid=(\d+)/i.exec(a.href);
      return m ? m[1] : null;
    }).filter(Boolean));
  }

  /* Build the same form post the browser sends when the Export button is clicked. */
  function formBody(form, btn) {
    var p = new URLSearchParams();
    Array.prototype.forEach.call(form.elements, function (el) {
      if (!el.name || el.disabled) return;
      var t = (el.type || '').toLowerCase();
      if (/^(submit|button|image|reset|file)$/.test(t)) return;
      if ((t === 'checkbox' || t === 'radio') && !el.checked) return;
      if (el.tagName === 'SELECT') {
        Array.prototype.forEach.call(el.options, function (o) { if (o.selected) p.append(el.name, o.value); });
        return;
      }
      p.append(el.name, el.value);
    });
    if (btn.name) {
      p.append(btn.name, btn.value || '');
    } else {
      var m = /__doPostBack\('([^']*)'/.exec(btn.getAttribute('onclick') || '');
      if (m) p.set('__EVENTTARGET', m[1]);
    }
    if (/multipart/i.test(form.getAttribute('enctype') || '')) {
      var fd = new FormData();
      p.forEach(function (v, k) { fd.append(k, v); });
      return fd;
    }
    return p;
  }

  function fileName(disposition, type, id) {
    var name = '';
    var m = /filename\*=UTF-8''([^;]+)/i.exec(disposition);
    if (m) name = decodeURIComponent(m[1]);
    else if ((m = /filename="?([^";]+)"?/i.exec(disposition))) name = m[1];
    if (!name) {
      var ext = /sheet|excel/i.test(type) ? 'xlsx' : /csv/i.test(type) ? 'csv' : /pdf/i.test(type) ? 'pdf' : 'dat';
      name = 'Transactions.' + ext;
    }
    return name.indexOf(id) >= 0 ? name : id + '_' + name;
  }

  function saveBlob(blob, name) {
    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(function () { URL.revokeObjectURL(a.href); }, 60000);
  }

  function isLogin(url) { return /\/Login/i.test(url); }

  function clean(s) { return (s || '').replace(/\s+/g, ' ').trim(); }

  function safeName(s) { return (s || '').replace(/[\\\/:*?"<>|\r\n\t]+/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 60); }

  /* Value shown next to a label such as "Customer Name:" - on the same line, or in a text box /
     text in the label's own row. preferBox checks the row first (requirement fields, whose label
     sits beside help text). Hidden elements are ignored, and the search never climbs into an
     area holding a second label, so it cannot pick up the next field's value. */
  function readField(doc, label, preferBox) {
    /* Label must be followed by ":" or end of line, so "State" never matches "Statement". */
    var match = new RegExp('^\\s*' + label + '\\s*(:|\\n|$)', 'i');
    var rx = new RegExp('^\\s*' + label + '\\s*:?\\s*', 'i');
    var shown = function (el) { return el.getClientRects().length > 0; };
    var els = Array.prototype.filter.call(doc.querySelectorAll('body *'), function (el) {
      if (/^(SCRIPT|STYLE|OPTION|SELECT|TEXTAREA)$/.test(el.tagName) || !shown(el)) return false;
      var t = (el.innerText || '').trim();
      return t.length < 400 && match.test(t);
    });
    els = els.filter(function (el) { return !els.some(function (o) { return o !== el && el.contains(o); }); });

    /* Labels for other fields inside an area (e.g. "City:" when looking for "Street Address"). */
    function otherLabels(area) {
      return Array.prototype.filter.call(area.querySelectorAll('*'), function (e) {
        var t = (e.innerText || '').trim();
        return !e.children.length && shown(e) && /^[^:\n]{2,60}:\s*$/.test(t) && !match.test(t);
      }).length;
    }
    function boxText(inp) {
      if (inp.tagName !== 'SELECT') return inp.value;
      var o = inp.options[inp.selectedIndex];
      return o && o.value !== '' && !/^\s*(--|select\b)/i.test(o.text) ? o.text : '';
    }
    function inlineValue() {
      for (var j = 0; j < els.length; j++) {
        var rest = els[j].innerText.trim().replace(rx, '').split('\n')[0].trim();
        if (rest) return rest;
      }
      return '';
    }
    var blank = false; /* the row has an empty box: the field is blank */
    function rowValue() {
      for (var j = 0; j < els.length; j++) {
        for (var a = els[j].parentElement, k = 0; a && k < 4; a = a.parentElement, k++) {
          if (otherLabels(a) > 0) break;
          var inp = a.querySelector('input:not([type=hidden]):not([type=button]):not([type=submit]):not([type=checkbox]):not([type=radio]), select, textarea');
          var val = inp ? boxText(inp) : '';
          if (val && val.trim()) return val.trim();
          if (inp) { blank = true; break; } /* do not read help text instead */
          var lines = (a.innerText || '').split('\n').map(function (x) { return x.trim(); }).filter(function (x) {
            return x && x !== '?' && !match.test(x) && !/:\s*$/.test(x);
          });
          if (lines.length) return lines[0];
        }
      }
      return '';
    }
    if (!preferBox) return inlineValue() || rowValue();
    var v = rowValue();
    return v || (blank ? '' : inlineValue());
  }

  function firstField(doc, labels) {
    for (var i = 0; i < labels.length; i++) {
      var v = readField(doc, labels[i], true);
      if (v) return v;
    }
    return '';
  }

  /* "US - United States" -> "US" for the combined address. */
  function shortCode(v) { var m = /^([A-Z]{2,3}) - /.exec(v); return m ? m[1] : v; }

  /* Address from the requirement fields: Street Address, City, State, Zip Code, Country
     (or the same fields prefixed "Permanent"). */
  function readAddress(doc, permanent) {
    var p = permanent ? 'Permanent ' : '';
    var a = {
      street: firstField(doc, permanent ? ['Permanent Address', 'Permanent Street Address'] : ['Street Address']),
      city: firstField(doc, [p + 'City']),
      state: firstField(doc, [p + 'State']),
      zip: firstField(doc, [p + 'Zip Code', p + 'Zip', p + 'Postal Code']),
      country: firstField(doc, [p + 'Country'])
    };
    a.full = [a.street, a.city, [shortCode(a.state), a.zip].filter(Boolean).join(' '), shortCode(a.country)]
      .filter(Boolean).join(', ');
    return a;
  }

  /* Load a page in an off-screen frame (not a pop-up, so never blocked), wait until ready(doc)
     returns something, then call use(doc, thatValue). On failure use(null, reason).
     ready() is tried as soon as the page's own content is parsed - it does not wait for
     embedded frames and images unless it chooses to. */
  async function withFrame(url, ready, use) {
    var f = document.createElement('iframe');
    f.style.cssText = FRAME_CSS;
    f.src = url;
    document.body.appendChild(f);
    try {
      var d, r = null, waited = 0;
      while (!r && waited < PAGE_TIMEOUT_MS) {
        await sleep(POLL_MS);
        waited += POLL_MS;
        try { d = f.contentDocument; } catch (e) { d = null; }
        if (!d) return await use(null, 'page could not load (error page or frames not allowed)');
        if (isLogin(d.location.href)) throw new Error('LOGIN');
        if (d.readyState !== 'loading' && d.location.href !== 'about:blank') r = ready(d);
      }
      return await use(r ? d : null, r || 'page did not finish loading');
    } finally {
      f.remove();
    }
  }

  /* Every stylesheet has arrived, so hidden help text is really hidden. */
  function stylesLoaded(d) {
    return Array.prototype.every.call(d.querySelectorAll('link[rel~="stylesheet"]'), function (l) { return !!l.sheet; });
  }

  async function waitComplete(d) {
    for (var waited = 0; d.readyState !== 'complete' && waited < PAGE_TIMEOUT_MS; waited += POLL_MS) await sleep(POLL_MS);
  }

  /* The Requirements page as the server sends it, fetched once per alert: the details and the
     narrative are both read from it. workAlert drops it when the alert is done. */
  var reqPages = {};
  function reqPage(id) {
    if (!reqPages[id]) {
      reqPages[id] = fetch(REQ_BASE + id, { credentials: 'same-origin' }).then(function (res) {
        if (isLogin(res.url)) throw new Error('LOGIN');
        return res.ok ? res.text() : null;
      }).then(function (html) {
        return html == null ? null : new DOMParser().parseFromString(html, 'text/html');
      });
    }
    return reqPages[id];
  }

  /* Fast path: fetch the Requirements page as plain HTML (no scripts, styles or frames)
     and read the fields by ECM's own element ids. Returns null when something needed is
     not in the raw HTML - the caller then falls back to the full-page readInfoFrame. */
  async function readInfoFast(id) {
    var d = await reqPage(id);
    if (!d) return null;

    /* Header lines written as "Label: value" (Customer Name, Customer Internal ID). */
    function header(label) {
      var rx = new RegExp('^\\s*' + label + '\\s*:\\s*([^\\n]*)', 'i'), best = null;
      Array.prototype.forEach.call(d.querySelectorAll('div,span,td,th,label,b,strong,p,li'), function (el) {
        var t = el.textContent.replace(/\s+/g, ' ').trim();
        if (t.length < 200 && rx.test(t)) best = el;   /* last match = innermost element */
      });
      if (!best) return '';
      return clean(best.textContent.replace(/\s+/g, ' ').trim().replace(new RegExp('^\\s*' + label + '\\s*:\\s*', 'i'), ''));
    }
    /* Requirement rows: <... id="..._ReqLabel_N">Label:</...> and the value in "..._divSoureField_N". */
    function req(label) {
      var rx = new RegExp('^\\s*' + label + '\\s*:?\\s*\\*?\\s*$', 'i');
      var labels = d.querySelectorAll('[id*="ReqLabel_"]');
      for (var i = 0; i < labels.length; i++) {
        if (!rx.test(labels[i].textContent)) continue;
        var m = /^(.*)ReqLabel_(\d+)$/.exec(labels[i].id);
        if (!m) continue;
        var field = d.getElementById(m[1] + 'divSoureField_' + m[2]);
        if (!field) return '';
        var inp = field.querySelector('input:not([type=hidden]):not([type=button]):not([type=submit]):not([type=checkbox]):not([type=radio]), select, textarea');
        if (inp && inp.tagName === 'SELECT') {
          var o = inp.querySelector('option[selected]');
          return o && o.value !== '' && !/^\s*(--|select\b)/i.test(o.textContent) ? clean(o.textContent) : '';
        }
        if (inp && inp.tagName === 'TEXTAREA') return clean(inp.textContent);
        if (inp) return clean(inp.getAttribute('value') || '');
        return clean(field.textContent);
      }
      return '';
    }
    function addr(permanent) {
      var p = permanent ? 'Permanent ' : '';
      var a = {
        street: req(permanent ? 'Permanent Address' : 'Street Address') || (permanent ? req('Permanent Street Address') : ''),
        city: req(p + 'City'), state: req(p + 'State'),
        zip: req(p + 'Zip Code') || req(p + 'Zip') || req(p + 'Postal Code'),
        country: req(p + 'Country')
      };
      a.full = [a.street, a.city, [shortCode(a.state), a.zip].filter(Boolean).join(' '), shortCode(a.country)].filter(Boolean).join(', ');
      return a;
    }

    var info = {
      customer: header('Customer Name'),
      internalId: header('Customer Internal ID'),
      velocity: req('Velocity Alert ID'),
      address: addr(false),
      permanent: addr(true)
    };
    return info.customer && info.velocity ? info : null;
  }

  /* Focal Entity info page: "Customer Details" rows, each a label
     (..._rptCustomerDetails_lblText_N, e.g. "ProgramDescription") and its value panel
     (..._rptCustomerDetails_pnlControl<anything>_N). Fields are found by label text, not
     by row number, since rows can differ between cases. Read from the raw HTML. */
  async function readFocal(id) {
    for (var b = 0; b < FOCAL_BASES.length; b++) {
      var res = await fetch(FOCAL_BASES[b] + id, { credentials: 'same-origin' });
      if (isLogin(res.url)) throw new Error('LOGIN');
      if (!res.ok) continue;
      var d = new DOMParser().parseFromString(await res.text(), 'text/html');
      if (!d.querySelector('[id*="rptCustomerDetails_lblText_"]')) continue;
      return {
        programId: focalField(d, 'ProgramId'),
        programName: focalField(d, 'ProgramName'),
        programDescription: focalField(d, 'ProgramDescription')
      };
    }
    return { error: 'Focal Entity info page not found' };
  }

  function focalField(d, label) {
    var norm = function (t) { return (t || '').replace(/[\s:_]+/g, '').toLowerCase(); };
    var labels = d.querySelectorAll('[id*="rptCustomerDetails_lblText_"]');
    for (var i = 0; i < labels.length; i++) {
      if (norm(labels[i].textContent) !== norm(label)) continue;
      var n = /_(\d+)$/.exec(labels[i].id);
      if (!n) continue;
      var panel = null;
      Array.prototype.some.call(d.querySelectorAll('[id*="rptCustomerDetails_pnlControl"]'), function (p) {
        if (new RegExp('_' + n[1] + '$').test(p.id)) { panel = p; return true; }
        return false;
      });
      if (!panel) return '';
      var inp = panel.querySelector('input:not([type=hidden]):not([type=button]):not([type=submit]), select, textarea');
      if (inp && inp.tagName === 'SELECT') {
        var o = inp.querySelector('option[selected]');
        return o ? clean(o.textContent) : '';
      }
      if (inp && inp.tagName === 'TEXTAREA') return clean(inp.textContent);
      if (inp) return clean(inp.getAttribute('value') || '');
      /* Text panel: "ProgramDescription: value" - drop the label part. */
      var t = clean(panel.textContent);
      var m = new RegExp('^\\s*' + label.split('').join('\\s*') + '\\s*:?\\s*', 'i').exec(t);
      return clean(m ? t.slice(m[0].length) : t);
    }
    return '';
  }

  async function readInfo(id) {
    var fast = await readInfoFast(id);
    if (fast) { fast.how = 'fast'; return fast; }
    var full = await readInfoFrame(id);
    if (full) full.how = 'full page';
    return full;
  }

  function readInfoFrame(id) {
    return withFrame(REQ_BASE + id, function (d) {
      return d.body && stylesLoaded(d) && /Customer Name/i.test(d.body.innerText) ? true : null;
    }, function (d, why) {
      if (!d) return { error: why };
      return {
        customer: readField(d, 'Customer Name'),
        internalId: readField(d, 'Customer Internal ID'),
        address: readAddress(d, false),
        permanent: readAddress(d, true),
        velocity: readField(d, 'Velocity Alert ID', true)
      };
    });
  }

  /* ---- Alert Analyst Decision and Alert Narrative (Requirements page, Alert Disposition tab) */
  var BOX = 'textarea, select, input:not([type=hidden]):not([type=button]):not([type=submit])' +
    ':not([type=checkbox]):not([type=radio])';

  /* The boxes holding a Requirements field's value, found by its label. On the Alert Disposition
     tab the box is not inside "..._divSoureField_N" but beside it, with an id that starts like the
     label's and ends in the same _N: "..._ReqLabel_2" -> "..._txtMulti<number>_2" (text) or
     "..._ddl<number>_0" (dropdown). Failing that, the nearest box around the label. Help-text
     copies of the labels ("..._FaqLabel_N") don't count. */
  function reqBoxes(d, label) {
    var rx = new RegExp('^\\s*\\*?\\s*' + label + '\\s*:?\\s*\\*?\\s*$', 'i'), found = [];
    var boxes = d.querySelectorAll(BOX);
    Array.prototype.forEach.call(d.querySelectorAll('[id*="ReqLabel_"]'), function (lab) {
      var m = /^(.*)ReqLabel_(\d+)$/.exec(lab.id);
      if (!m || !rx.test(lab.textContent)) return;
      var box = null, end = '_' + m[2];
      Array.prototype.some.call(boxes, function (el) {
        var i = el.id || '';
        if (i.indexOf(m[1]) === 0 && i.slice(-end.length) === end) { box = el; return true; }
        return false;
      });
      for (var a = lab.parentElement, k = 0; !box && a && k < 4; a = a.parentElement, k++) box = a.querySelector(BOX);
      if (box) found.push(box);
    });
    return found;
  }

  /* A box's value: the chosen option of a dropdown, or the text with its line breaks kept.
     live: a page in a frame (read what is showing) rather than the HTML as sent. */
  function boxValue(box, live) {
    if (box.tagName === 'SELECT') {
      var o = live ? box.options[box.selectedIndex] : box.querySelector('option[selected]');
      return o && o.value !== '' && !/^\s*(--|select\b)/i.test(o.textContent) ? clean(o.textContent) : '';
    }
    var t = live ? box.value : box.tagName === 'TEXTAREA' ? box.textContent : box.getAttribute('value');
    return (t || '').replace(/\r\n?/g, '\n').replace(/[ \t]+$/gm, '').trim();
  }

  /* Both fields from one page; null when neither label is on it (that tab isn't loaded). */
  function narrativeFrom(d, live) {
    var pick = function (boxes) {   /* the first box that says something */
      for (var i = 0; i < boxes.length; i++) { var v = boxValue(boxes[i], live); if (v) return v; }
      return '';
    };
    var nb = reqBoxes(d, 'Alert Narrative'), db = reqBoxes(d, 'Alert Analyst Decision');
    return nb.length || db.length ? { narrative: pick(nb), decision: pick(db) } : null;
  }

  /* From the Requirements page as sent (fast); when the narrative isn't there, from the page in a
     frame with its Alert Disposition tab opened. */
  async function readNarrative(id) {
    var d = await reqPage(id);
    var n = d ? narrativeFrom(d, false) : null;
    if (n && n.narrative) { n.how = 'fast'; return n; }
    var full = await readNarrativeFrame(id);
    if (full && !full.error) { full.how = 'full page'; return full; }
    if (n) { n.how = 'fast'; return n; }   /* the fields are there, the narrative is empty */
    return full;
  }

  function readNarrativeFrame(id) {
    var clicked = false;
    return withFrame(REQ_BASE + id, function (d) {
      if (!d.body) return null;
      if (reqBoxes(d, 'Alert Narrative').length) return true;
      if (!clicked && d.readyState === 'complete') {
        var tabs = Array.prototype.filter.call(d.querySelectorAll('a'), function (a) {
          return clean(a.textContent) === 'Alert Disposition';
        });
        var tab = tabs.filter(function (a) { return /requirement-font-size/.test(a.className); })[0] || tabs[0];
        if (tab) { clicked = true; tab.click(); }
      }
      return null;
    }, function (d, why) {
      if (!d) return { error: why === 'page did not finish loading' ? 'Alert Narrative not found on the page' : why };
      return narrativeFrom(d, true) || { error: 'Alert Narrative not found on the page' };
    });
  }

  /* Post the live page's form ourselves so the file can be named and confirmed.
     Returns { blob, disp, type } or null when the server sent no file. */
  async function postExport(d, btn) {
    var form = btn.form || btn.closest('form');
    if (!form) return null;
    var action = new URL(form.getAttribute('action') || d.location.href, d.location.href).href;
    var post = await fetch(action, { method: 'POST', credentials: 'same-origin', body: formBody(form, btn) });
    if (isLogin(post.url)) throw new Error('LOGIN');
    var type = post.headers.get('content-type') || '';
    var disp = post.headers.get('content-disposition') || '';
    if (!post.ok || (!/attachment/i.test(disp) && /text\/html/i.test(type))) return null;
    var blob = await post.blob();
    /* An error page sent without a proper content type is not a file. */
    if (!/attachment/i.test(disp) && /^\s*<(!doctype|html)/i.test(await blob.slice(0, 200).text())) return null;
    return { blob: blob, disp: disp, type: type };
  }

  /* Export from the Statistics page. Tried as soon as the button exists; if that gets no
     file, again once the whole page (embedded frames too) has loaded; last of all the real
     button is clicked. Returns { blob, disp, type } or { result: text }. */
  function getExport(id) {
    return withFrame(BASE + id, function (d) { return d.getElementById(BUTTON_ID); }, async function (d, btnEl) {
      if (!d) return { result: btnEl === 'page did not finish loading' ? 'no Export button on page' : btnEl };
      var file = await postExport(d, btnEl);
      if (file) return file;
      await waitComplete(d);
      file = await postExport(d, btnEl);
      if (file) return file;
      btnEl.click();
      await sleep(CLICK_WAIT_MS);
      return { result: 'OK (clicked Export, file keeps its original name)' };
    });
  }

  function csvCell(v) {
    v = String(v == null ? '' : v);
    if (/^[=+\-@]/.test(v)) v = "'" + v; /* stop Excel treating a value as a formula */
    return '"' + v.replace(/"/g, '""') + '"';
  }

  function tableRows() {
    return results.map(function (r) {
      var a = r.address || {};
      return [r.ecm, r.velocity, r.customer, r.internalId, a.full, a.street, a.city, a.state, a.zip, a.country,
        (r.permanent || {}).full, r.programId, r.programName, r.programDescription, r.file, r.result];
    });
  }

  /* yyyymmdd_hhmm, for the file names */
  function stampNow() {
    var now = new Date();
    return now.getFullYear() + ('0' + (now.getMonth() + 1)).slice(-2) + ('0' + now.getDate()).slice(-2) + '_' +
      ('0' + now.getHours()).slice(-2) + ('0' + now.getMinutes()).slice(-2);
  }

  function saveCsv(stamp) {
    if (!results.length) return;
    var lines = [COLUMNS].concat(tableRows()).map(function (cells) { return cells.map(csvCell).join(','); });
    saveBlob(new Blob(['\ufeff' + lines.join('\r\n')], { type: 'text/csv' }), 'ECM_Alert_Info_' + stamp + '.csv');
  }

  /* ---- ECM_Narratives_<date>_<time>.xlsx: one row per alert, in the order the IDs were given.
     An .xlsx is a zip of XML files; this writes the few Excel needs, stored without compression. */
  var CELL_MAX = 32767;   /* the most characters Excel holds in one cell */

  /* A narrative as cells of at most CELL_MAX characters. Only a longer one is split, at the last
     space or line break before the limit, so no word is cut; the next cell starts at the next word. */
  function splitCells(text) {
    var parts = [];
    while (text.length > CELL_MAX) {
      var head = text.slice(0, CELL_MAX + 1);
      var cut = Math.max(head.lastIndexOf(' '), head.lastIndexOf('\n'), head.lastIndexOf('\t'));
      if (cut < CELL_MAX / 2) {   /* no break anywhere near the limit: cut at it */
        cut = CELL_MAX;
        var c = text.charCodeAt(cut - 1);
        if (c >= 0xD800 && c <= 0xDBFF) cut--;   /* not between the two halves of one character */
      }
      parts.push(text.slice(0, cut).replace(/\s+$/, ''));
      text = text.slice(cut).replace(/^\s+/, '');
    }
    parts.push(text);
    return parts;
  }

  function xmlText(s) {
    return String(s == null ? '' : s).replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\ufffe\uffff]/g, '')
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  /* 0 -> A, 25 -> Z, 26 -> AA */
  function colName(i) {
    var s = '';
    for (i++; i > 0; i = Math.floor((i - 1) / 26)) s = String.fromCharCode(65 + (i - 1) % 26) + s;
    return s;
  }

  /* Text is written into the cell itself, so a value starting with "=" stays text. */
  function cellXml(ref, v, style) {
    if (typeof v === 'number') return '<c r="' + ref + '" s="' + style + '"><v>' + v + '</v></c>';
    return '<c r="' + ref + '" s="' + style + '" t="inlineStr"><is><t xml:space="preserve">' + xmlText(v) + '</t></is></c>';
  }

  /* The sheet: header row frozen and filtered; Narrative columns wide, wrapped and top-aligned.
     "Narrative (continued)" columns are added only when a narrative in this run needs them.
     Styles: 1 header, 2 wrapped text, 3 other cells. */
  function narrativeSheet() {
    var rows = results.map(function (r, i) {
      var parts = r.narrative ? splitCells(r.narrative) : [''];
      var res = r.narrError ? 'error: ' + r.narrError : !r.narrative ? 'no narrative found' :
        parts.length > 1 ? 'OK, split into ' + parts.length + ' cells' : 'OK';
      var ecm = /^\d{1,15}$/.test(r.ecm) ? +r.ecm : r.ecm;
      return { vals: [i + 1, ecm, r.velocity, r.customer, r.decision], parts: parts, result: res };
    });
    var nCont = Math.max.apply(null, rows.map(function (r) { return r.parts.length; })) - 1;
    var head = ['#', 'ECM ID', 'Velocity Alert ID', 'Customer Name', 'Alert Analyst Decision', 'Narrative'];
    var widths = [5, 10, 18, 28, 24, 100];
    for (var k = 1; k <= nCont; k++) {
      head.push(nCont > 1 ? 'Narrative (continued ' + k + ')' : 'Narrative (continued)');
      widths.push(100);
    }
    head.push('Result');
    widths.push(26);
    var ref = 'A1:' + colName(head.length - 1) + (rows.length + 1);
    var x = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
      '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>',
      '</sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/><cols>'];
    widths.forEach(function (w, i) { x.push('<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + w + '" customWidth="1"/>'); });
    x.push('</cols><sheetData><row r="1">');
    head.forEach(function (h, i) { x.push(cellXml(colName(i) + '1', h, 1)); });
    x.push('</row>');
    rows.forEach(function (r, j) {
      var n = j + 2, vals = r.vals.concat(r.parts);
      while (vals.length < 6 + nCont) vals.push('');
      vals.push(r.result);
      x.push('<row r="' + n + '">');
      vals.forEach(function (v, i) {
        if (v !== '' && v != null) x.push(cellXml(colName(i) + n, v, i >= 5 && i < 6 + nCont ? 2 : 3));
      });
      x.push('</row>');
    });
    x.push('</sheetData><autoFilter ref="' + ref + '"/></worksheet>');
    return { xml: x.join(''), ref: ref };
  }

  function saveNarratives(stamp) {
    if (!results.length) return false;
    var sheet = narrativeSheet();
    var NS = 'http://schemas.openxmlformats.org/', HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
    var files = [
      ['[Content_Types].xml', HEAD + '<Types xmlns="' + NS + 'package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
        '</Types>'],
      ['_rels/.rels', HEAD + '<Relationships xmlns="' + NS + 'package/2006/relationships">' +
        '<Relationship Id="rId1" Type="' + NS + 'officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
        '</Relationships>'],
      ['xl/workbook.xml', HEAD + '<workbook xmlns="' + NS + 'spreadsheetml/2006/main" xmlns:r="' + NS +
        'officeDocument/2006/relationships"><sheets><sheet name="Narratives" sheetId="1" r:id="rId1"/></sheets>' +
        '<definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">Narratives!' +
        sheet.ref.replace(/([A-Z]+)(\d+)/g, '$$$1$$$2') + '</definedName></definedNames></workbook>'],
      ['xl/_rels/workbook.xml.rels', HEAD + '<Relationships xmlns="' + NS + 'package/2006/relationships">' +
        '<Relationship Id="rId1" Type="' + NS + 'officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
        '<Relationship Id="rId2" Type="' + NS + 'officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
        '</Relationships>'],
      ['xl/worksheets/sheet1.xml', sheet.xml],
      ['xl/styles.xml', HEAD + '<styleSheet xmlns="' + NS + 'spreadsheetml/2006/main">' +
        '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>' +
        '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
        '<fill><patternFill patternType="solid"><fgColor rgb="FFD9E1F2"/><bgColor indexed="64"/></patternFill></fill></fills>' +
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
        '<cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
        '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment vertical="center"/></xf>' +
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>' +
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top"/></xf></cellXfs>' +
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>']
    ];
    saveBlob(zipStored(files), 'ECM_Narratives_' + stamp + '.xlsx');
    return true;
  }

  /* A zip with each file stored as it is (no compression). */
  var CRC_TABLE = null;
  function crc32(bytes) {
    if (!CRC_TABLE) {
      CRC_TABLE = new Uint32Array(256);
      for (var n = 0; n < 256; n++) {
        var c = n;
        for (var k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
        CRC_TABLE[n] = c >>> 0;
      }
    }
    var crc = 0xFFFFFFFF;
    for (var i = 0; i < bytes.length; i++) crc = CRC_TABLE[(crc ^ bytes[i]) & 0xFF] ^ (crc >>> 8);
    return (crc ^ 0xFFFFFFFF) >>> 0;
  }

  function zipStored(files) {
    var enc = new TextEncoder(), parts = [], central = [], offset = 0, cdSize = 0;
    files.forEach(function (f) {
      var name = enc.encode(f[0]), data = enc.encode(f[1]), crc = crc32(data);
      var h = new DataView(new ArrayBuffer(30));
      h.setUint32(0, 0x04034b50, true); h.setUint16(4, 20, true); h.setUint16(12, 0x21, true);
      h.setUint32(14, crc, true); h.setUint32(18, data.length, true); h.setUint32(22, data.length, true);
      h.setUint16(26, name.length, true);
      parts.push(new Uint8Array(h.buffer), name, data);
      var c = new DataView(new ArrayBuffer(46));
      c.setUint32(0, 0x02014b50, true); c.setUint16(4, 20, true); c.setUint16(6, 20, true); c.setUint16(14, 0x21, true);
      c.setUint32(16, crc, true); c.setUint32(20, data.length, true); c.setUint32(24, data.length, true);
      c.setUint16(28, name.length, true); c.setUint32(42, offset, true);
      central.push(new Uint8Array(c.buffer), name);
      offset += 30 + name.length + data.length;
      cdSize += 46 + name.length;
    });
    var e = new DataView(new ArrayBuffer(22));
    e.setUint32(0, 0x06054b50, true); e.setUint16(8, files.length, true); e.setUint16(10, files.length, true);
    e.setUint32(12, cdSize, true); e.setUint32(16, offset, true);
    return new Blob(parts.concat(central, [new Uint8Array(e.buffer)]),
      { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }

  function copyForExcel() {
    if (!results.length) { log('Nothing to copy yet - press Start first.'); return; }
    var tsv = [COLUMNS].concat(tableRows()).map(function (cells) {
      return cells.map(function (c) { return String(c == null ? '' : c).replace(/[\t\r\n]+/g, ' '); }).join('\t');
    }).join('\r\n');
    var done = function () { log('Copied ' + results.length + ' rows. Click a cell in Excel and press Ctrl+V.'); };
    var fallback = function () {
      var ta = document.createElement('textarea');
      ta.value = tsv; document.body.appendChild(ta); ta.select();
      document.execCommand('copy'); ta.remove(); done();
    };
    if (navigator.clipboard) navigator.clipboard.writeText(tsv).then(done, fallback); else fallback();
  }

  /* ---------- panel ---------- */
  var panel = document.createElement('div');
  panel.id = '__ecmExport';
  /* Every style is set inline and in full, and class names are unique (the ECM page uses
     Bootstrap, whose .row adds invisible ::before/::after items that would sit in the grid),
     so the page's own CSS cannot shift the layout. */
  var FONT = 'font-family:Segoe UI,Arial,sans-serif;font-size:13px;line-height:1.4;letter-spacing:normal;text-transform:none;';
  panel.style.cssText = 'position:fixed;right:16px;bottom:16px;z-index:2147483647;width:360px;box-sizing:border-box;' +
    'background:#1f2937;color:#f9fafb;padding:12px;border-radius:10px;box-shadow:0 6px 24px rgba(0,0,0,.4);text-align:left;' + FONT;
  panel.innerHTML =
    '<div style="display:flex;align-items:center;justify-content:space-between;margin:0 0 8px;' + FONT + '">' +
      '<span style="font-weight:600;font-size:14px">ECM Export Transactions ' +
        '<span style="font-weight:400;font-size:11px;color:#9ca3af">v' + VERSION + '</span></span>' +
      '<span style="display:flex;align-items:center;gap:8px">' +
        '<span class="__ecmCount" style="font-size:12px;color:#9ca3af"></span>' +
        '<button class="__ecmClose" title="Close" style="all:unset;cursor:pointer;width:22px;height:22px;line-height:22px;text-align:center;' +
          'border-radius:4px;color:#d1d5db;font:16px Arial,sans-serif">&#x2715;</button>' +
      '</span>' +
    '</div>' +
    '<textarea rows="6" placeholder="Alert IDs - one per line, or separated by commas/spaces" style="display:block;width:100%;' +
      'box-sizing:border-box;margin:0;padding:6px 8px;border:1px solid #4b5563;border-radius:6px;background:#f9fafb;color:#111827;' +
      'font:13px/1.4 Consolas,monospace;resize:vertical"></textarea>' +
    '<div style="display:flex;align-items:center;gap:8px;margin:8px 0 0;' + FONT + 'font-size:12px;color:#d1d5db">' +
      '<span>Alerts at a time:</span>' +
      '<select class="__ecmPar" style="appearance:auto;-webkit-appearance:menulist;box-sizing:border-box;margin:0;' +
        'width:56px;height:26px;padding:0 4px;border:1px solid #4b5563;border-radius:5px;background:#f9fafb;color:#111827;' +
        'font:12px Segoe UI,Arial,sans-serif;line-height:normal;text-transform:none;cursor:pointer">' +
        '<option value="1">1</option><option value="2">2</option><option value="3">3</option></select>' +
      '<span style="color:#9ca3af">more = faster, more load on ECM</span>' +
    '</div>' +
    '<div class="__ecmRow" style="display:grid;grid-template-columns:1fr 1fr;gap:6px;margin:8px 0"></div>' +
    '<div class="__ecmLog" style="max-height:180px;overflow:auto;padding:6px 8px;border-radius:6px;background:#111827;color:#e5e7eb;' +
      'font:12px/1.45 Consolas,monospace;white-space:pre-wrap;min-height:18px"></div>';
  var box = panel.querySelector('textarea');
  var row = panel.querySelector('.__ecmRow');
  var logEl = panel.querySelector('.__ecmLog');
  var countEl = panel.querySelector('.__ecmCount');
  var parEl = panel.querySelector('.__ecmPar');
  parEl.value = String(Math.min(3, Math.max(1, PARALLEL)));
  parEl.onchange = function () {
    PARALLEL = +parEl.value;
    try { localStorage.setItem('__ecmParallel', String(PARALLEL)); } catch (e) { /* storage blocked */ }
  };

  function updateCount() {
    var n = parseIds(box.value).length;
    countEl.textContent = n ? n + (n === 1 ? ' ID' : ' IDs') : '';
  }
  box.addEventListener('input', updateCount);

  function btn(label, fn, primary, wide) {
    var b = document.createElement('button');
    b.textContent = label;
    b.style.cssText = 'all:unset;box-sizing:border-box;display:block;width:100%;height:30px;line-height:30px;text-align:center;' +
      (wide ? 'grid-column:1 / span 2;' : '') +
      'border-radius:6px;cursor:pointer;font:13px Segoe UI,Arial,sans-serif;font-weight:' + (primary ? '600' : '400') + ';' +
      (primary ? 'background:#2563eb;color:#fff;' : 'background:#e5e7eb;color:#111827;');
    b.onmouseenter = function () { b.style.filter = 'brightness(0.92)'; };
    b.onmouseleave = function () { b.style.filter = ''; };
    b.onclick = function (e) { e.preventDefault(); fn(); };
    row.appendChild(b);
    return b;
  }

  function log(msg) {
    var line = document.createElement('div');
    line.style.cssText = 'all:unset;display:block;font:inherit;line-height:inherit;color:inherit;white-space:pre-wrap';
    line.textContent = msg;
    logEl.appendChild(line);
    logEl.scrollTop = logEl.scrollHeight;
    return line;
  }

  function secs(ms) { return (ms / 1000).toFixed(1) + ' s'; }

  /* One alert: details and export pages side by side. Returns a results row. */
  async function workAlert(id, line, tagStart) {
    var tag = tagStart, info = {}, focal = {}, narr = {}, result, t0 = Date.now(), tInfo = 0, tExp = 0, tNarr = 0;
    try {
      var both = await Promise.all([
        readInfo(id).then(function (v) { tInfo = Date.now() - t0; return v; }),
        getExport(id).then(function (v) { tExp = Date.now() - t0; return v; }),
        readFocal(id).catch(function (e) { if (e.message === 'LOGIN') throw e; return { error: e.message }; }),
        readNarrative(id).then(function (v) { tNarr = Date.now() - t0; return v; },
          function (e) { if (e.message === 'LOGIN') throw e; return { error: e.message }; })
      ]);
      info = both[0];
      var ex = both[1];
      focal = both[2] || {};
      narr = both[3] || {};
      if (info.error) tag += ' (details: ' + info.error + ')';
      else tag += ' ' + (info.velocity || '(no Velocity ID found)');
      if (ex.blob) {
        var name = fileName(ex.disp, ex.type, info.velocity ? id + '_' + safeName(info.velocity) : id);
        saveBlob(ex.blob, name);
        result = 'OK ' + name;
      } else {
        result = ex.result;
      }
    } catch (e) {
      if (e.message === 'LOGIN') throw e;
      result = 'error: ' + e.message;
    } finally {
      delete reqPages[id];
    }
    line.textContent = tag + ' ' + result + '  (' + secs(Date.now() - t0) +
      (tInfo && tExp ? ': details ' + secs(tInfo) + (info.how ? ' (' + info.how + ')' : '') + ', export ' + secs(tExp) : '') +
      (narr.error ? ', narrative: ' + narr.error : tNarr ? ', narrative ' + secs(tNarr) + (narr.how ? ' (' + narr.how + ')' : '') +
        (narr.narrative ? '' : ' - empty') : '') + ')';
    return {
      ecm: id, velocity: info.velocity || '', customer: info.customer || '', internalId: info.internalId || '',
      address: info.address, permanent: info.permanent,
      programId: focal.programId || '', programName: focal.programName || '', programDescription: focal.programDescription || '',
      decision: narr.decision || '', narrative: narr.narrative || '', narrError: narr.error || '',
      file: result.indexOf('OK ') === 0 && result.indexOf('OK (') !== 0 ? result.slice(3) : '', result: result
    };
  }

  async function run() {
    if (running) return;
    var ids = parseIds(box.value);
    if (!ids.length) { log('Enter at least one alert ID.'); return; }
    running = true; stopped = false;
    logEl.textContent = '';
    var rows = new Array(ids.length), next = 0, loggedOut = false, t0 = Date.now();

    /* PARALLEL workers each take the next alert until none are left. */
    async function worker() {
      while (!stopped && !loggedOut && next < ids.length) {
        var i = next++;
        var tag = '[' + (i + 1) + '/' + ids.length + '] ' + ids[i];
        var line = log(tag + ' working ...');
        try {
          rows[i] = await workAlert(ids[i], line, tag);
        } catch (e) {
          loggedOut = true;
          line.textContent = tag + ' session expired';
          return;
        }
        await sleep(PAUSE_MS);
      }
    }
    var workers = [];
    for (var w = 0; w < Math.max(1, PARALLEL); w++) workers.push(worker());
    await Promise.all(workers);

    results = rows.filter(Boolean);
    var ok = results.filter(function (r) { return r.result.indexOf('OK') === 0; }).length;
    var failed = results.filter(function (r) { return r.result.indexOf('OK') !== 0; }).map(function (r) { return r.ecm; });
    if (loggedOut) log('Log in to ECM again, then press Start.');
    log((stopped || loggedOut ? 'Stopped. ' : 'Finished. ') + ok + ' exported' +
        (failed.length ? ', failed: ' + failed.join(', ') : '.') + '  Total ' + secs(Date.now() - t0) + '.');
    var stamp = stampNow();
    saveCsv(stamp);
    if (results.length) log('Saved ECM_Alert_Info CSV to Downloads. "Copy for Excel" copies the same table.');
    try {
      if (saveNarratives(stamp)) log('Saved ECM_Narratives_' + stamp + '.xlsx to Downloads.');
    } catch (e) {
      log('Could not save the narratives file: ' + e.message);
    }
    running = false;
  }

  btn('Start', run, true);
  btn('Stop', function () { stopped = true; });
  btn('Load my queue', function () {
    var ids = queueIds();
    box.value = ids.join('\n');
    updateCount();
    log(ids.length ? 'Loaded ' + ids.length + ' IDs from My Case Queue.' : 'No queue found - open the Dashboard, or type IDs.');
  });
  btn('Clear', function () {
    if (running) { log('Export is running - press Stop first.'); return; }
    box.value = '';
    logEl.textContent = '';
    updateCount();
    box.focus();
  });
  btn('Copy for Excel', copyForExcel, false, true);
  panel.querySelector('.__ecmClose').onclick = function (e) { e.preventDefault(); stopped = true; panel.remove(); };

  document.body.appendChild(panel);
  var initial = queueIds();
  if (initial.length) { box.value = initial.join('\n'); updateCount(); log('Loaded ' + initial.length + ' IDs from My Case Queue. Remove any you do not want, then press Start.'); }
  box.focus();
})();
