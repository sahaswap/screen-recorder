/*
  ECM Transactions Export
  For each alert ID (ECM case ID):
    1. reads Customer Name, Customer Internal ID and Velocity Alert ID from the Requirements page;
    2. runs "Export Transactions" on the Statistics page, saving the file to Downloads as
       <ECM ID>_<Velocity Alert ID>_<original file name>;
  then saves ECM_Alert_Info_<date>.csv (one row per alert) for the Excel tool.
  Run on any ECM page while logged in: paste into the Console (Ctrl+Shift+J) or use the bookmarklet.
  Uses your existing Edge session - no passwords are stored or sent anywhere else.
*/
(function () {
  if (document.getElementById('__ecmExport')) return;

  var BASE = location.origin + '/ECM/Case/Statistics.aspx?caseid=';
  var REQ_BASE = location.origin + '/ECM/Case/Requirements.aspx?caseid=';
  /* Off-screen but rendered: page text is only readable from a frame that is not hidden. */
  var FRAME_CSS = 'position:fixed;left:-10000px;top:0;width:1200px;height:800px;border:0;opacity:0;pointer-events:none';
  var COLUMNS = ['ECM ID', 'Velocity Alert ID', 'Customer Name', 'Customer Internal ID', 'Transaction File', 'Result'];
  var results = [];
  var BUTTON_ID = 'btnExportStatsTransactions';
  var CLICK_WAIT_MS = 10000; /* after clicking Export: time allowed for the download to start */
  var PAUSE_MS = 1000;     /* pause between alerts, to go easy on the server */
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

  function safeName(s) { return (s || '').replace(/[\\\/:*?"<>|\r\n\t]+/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 60); }

  /* Value shown next to a label such as "Customer Name:" - on the same line, or in a text box /
     text in the label's own row. preferBox checks the row first (requirement fields, whose label
     sits beside help text). Hidden elements are ignored, and the search never climbs into an
     area holding a second label, so it cannot pick up the next field's value. */
  function readField(doc, label, preferBox) {
    var rx = new RegExp('^\\s*' + label + '\\s*:?\\s*', 'i');
    var shown = function (el) { return el.getClientRects().length > 0; };
    var els = Array.prototype.filter.call(doc.querySelectorAll('body *'), function (el) {
      if (/^(SCRIPT|STYLE|OPTION|SELECT|TEXTAREA)$/.test(el.tagName) || !shown(el)) return false;
      var t = (el.innerText || '').trim();
      return t.length < 400 && rx.test(t);
    });
    els = els.filter(function (el) { return !els.some(function (o) { return o !== el && el.contains(o); }); });

    function labelCount(area) {
      return Array.prototype.filter.call(area.querySelectorAll('*'), function (e) {
        return !e.children.length && shown(e) && /^[^:\n]{2,60}:\s*$/.test((e.innerText || '').trim());
      }).length;
    }
    function inlineValue() {
      for (var j = 0; j < els.length; j++) {
        var rest = els[j].innerText.trim().replace(rx, '').split('\n')[0].trim();
        if (rest) return rest;
      }
      return '';
    }
    function rowValue() {
      for (var j = 0; j < els.length; j++) {
        for (var a = els[j].parentElement, k = 0; a && k < 4; a = a.parentElement, k++) {
          if (labelCount(a) > 1) break;
          var inp = a.querySelector('input:not([type=hidden]):not([type=button]):not([type=submit]):not([type=checkbox]):not([type=radio]), select, textarea');
          var val = inp ? (inp.tagName === 'SELECT' ? (inp.options[inp.selectedIndex] || {}).text : inp.value) : '';
          if (val && val.trim()) return val.trim();
          var lines = (a.innerText || '').split('\n').map(function (x) { return x.trim(); }).filter(function (x) {
            return x && x !== '?' && !rx.test(x) && !/:\s*$/.test(x);
          });
          if (lines.length) return lines[0];
        }
      }
      return '';
    }
    return preferBox ? (rowValue() || inlineValue()) : (inlineValue() || rowValue());
  }

  /* Load a page in an off-screen frame (not a pop-up, so never blocked), wait until ready(doc)
     returns something, then call use(doc, thatValue). On failure use(null, reason). */
  async function withFrame(url, ready, use) {
    var f = document.createElement('iframe');
    f.style.cssText = FRAME_CSS;
    f.src = url;
    document.body.appendChild(f);
    try {
      var d, r = null;
      for (var i = 0; i < 120 && !r; i++) {
        await sleep(500);
        try { d = f.contentDocument; } catch (e) { d = null; }
        if (!d) return await use(null, 'page refuses to load in a frame');
        if (isLogin(d.location.href)) throw new Error('LOGIN');
        if (d.readyState === 'complete' && d.location.href !== 'about:blank') r = ready(d);
      }
      return await use(r ? d : null, r || 'page did not finish loading');
    } finally {
      f.remove();
    }
  }

  function readInfo(id) {
    return withFrame(REQ_BASE + id, function (d) {
      return d.body && /Customer Name/i.test(d.body.innerText) ? true : null;
    }, function (d, why) {
      if (!d) return { error: why };
      return {
        customer: readField(d, 'Customer Name'),
        internalId: readField(d, 'Customer Internal ID'),
        velocity: readField(d, 'Velocity Alert ID', true)
      };
    });
  }

  /* Post the live page's form ourselves so the file can be named and confirmed. */
  async function postExport(d, btn, id) {
    var form = btn.form || btn.closest('form');
    if (!form) return 'no form';
    var action = new URL(form.getAttribute('action') || d.location.href, d.location.href).href;
    var post = await fetch(action, { method: 'POST', credentials: 'same-origin', body: formBody(form, btn) });
    if (isLogin(post.url)) throw new Error('LOGIN');
    var type = post.headers.get('content-type') || '';
    var disp = post.headers.get('content-disposition') || '';
    if (!post.ok || (!/attachment/i.test(disp) && /text\/html/i.test(type))) return 'no file';
    var blob = await post.blob();
    /* An error page sent without a proper content type is not a file. */
    if (!/attachment/i.test(disp) && /^\s*<(!doctype|html)/i.test(await blob.slice(0, 200).text())) return 'no file';
    var name = fileName(disp, type, id);
    saveBlob(blob, name);
    return 'OK ' + name;
  }

  /* Export from the Statistics page. If our own post gets no file, click the real button instead. */
  function exportAlert(id, prefix) {
    return withFrame(BASE + id, function (d) { return d.getElementById(BUTTON_ID); }, async function (d, btnEl) {
      if (!d) return btnEl === 'page did not finish loading' ? 'no Export button on page' : btnEl;
      var result = await postExport(d, btnEl, prefix);
      if (result.indexOf('OK') === 0) return result;
      btnEl.click();
      await sleep(CLICK_WAIT_MS);
      return 'OK (clicked Export, file keeps its original name)';
    });
  }

  function csvCell(v) {
    v = String(v == null ? '' : v);
    if (/^[=+\-@]/.test(v)) v = "'" + v; /* stop Excel treating a value as a formula */
    return '"' + v.replace(/"/g, '""') + '"';
  }

  function tableRows() {
    return results.map(function (r) { return [r.ecm, r.velocity, r.customer, r.internalId, r.file, r.result]; });
  }

  function saveCsv() {
    if (!results.length) return;
    var lines = [COLUMNS].concat(tableRows()).map(function (cells) { return cells.map(csvCell).join(','); });
    var now = new Date();
    var stamp = now.getFullYear() + ('0' + (now.getMonth() + 1)).slice(-2) + ('0' + now.getDate()).slice(-2) + '_' +
      ('0' + now.getHours()).slice(-2) + ('0' + now.getMinutes()).slice(-2);
    saveBlob(new Blob(['\ufeff' + lines.join('\r\n')], { type: 'text/csv' }), 'ECM_Alert_Info_' + stamp + '.csv');
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
  panel.style.cssText = 'position:fixed;right:12px;bottom:12px;z-index:2147483647;width:340px;background:#1f2937;color:#fff;' +
    'font:13px/1.4 Segoe UI,Arial,sans-serif;padding:10px 12px;border-radius:8px;box-shadow:0 4px 16px rgba(0,0,0,.35)';
  panel.innerHTML =
    '<div style="font-weight:600;margin-bottom:6px">ECM Export Transactions</div>' +
    '<textarea rows="6" style="width:100%;box-sizing:border-box;font:12px Consolas,monospace;color:#111" ' +
    'placeholder="Alert IDs - one per line, or separated by commas/spaces"></textarea>' +
    '<div class="row" style="margin:6px 0"></div>' +
    '<div class="log" style="max-height:180px;overflow:auto;font:12px Consolas,monospace;white-space:pre-wrap"></div>';
  var box = panel.querySelector('textarea');
  var row = panel.querySelector('.row');
  var logEl = panel.querySelector('.log');

  function btn(label, fn) {
    var b = document.createElement('button');
    b.textContent = label;
    b.style.cssText = 'margin-right:4px;padding:3px 8px;border:0;border-radius:4px;cursor:pointer;background:#e5e7eb;color:#111';
    b.onclick = function (e) { e.preventDefault(); fn(); };
    row.appendChild(b);
    return b;
  }

  function log(msg) {
    var line = document.createElement('div');
    line.textContent = msg;
    logEl.appendChild(line);
    logEl.scrollTop = logEl.scrollHeight;
    return line;
  }

  async function run() {
    if (running) return;
    var ids = parseIds(box.value);
    if (!ids.length) { log('Enter at least one alert ID.'); return; }
    running = true; stopped = false;
    logEl.textContent = '';
    results = [];
    var ok = 0, failed = [];
    for (var i = 0; i < ids.length && !stopped; i++) {
      var id = ids[i];
      var tag = '[' + (i + 1) + '/' + ids.length + '] ' + id;
      var line = log(tag + ' reading details ...');
      var info = {}, result;
      try {
        info = await readInfo(id);
        if (info.error) tag += ' (details: ' + info.error + ')';
        else tag += ' ' + (info.velocity || '(no Velocity ID found)');
        line.textContent = tag + ' exporting ...';
        var prefix = info.velocity ? id + '_' + safeName(info.velocity) : id;
        result = await exportAlert(id, prefix);
      } catch (e) {
        if (e.message === 'LOGIN') { line.textContent = tag + ' session expired'; log('Log in to ECM again, then press Start.'); break; }
        result = 'error: ' + e.message;
      }
      line.textContent = tag + ' ' + result;
      results.push({
        ecm: id, velocity: info.velocity || '', customer: info.customer || '', internalId: info.internalId || '',
        file: result.indexOf('OK ') === 0 && result.indexOf('OK (') !== 0 ? result.slice(3) : '', result: result
      });
      if (result.indexOf('OK') === 0) ok++; else failed.push(id);
      await sleep(PAUSE_MS);
    }
    log((stopped ? 'Stopped. ' : 'Finished. ') + ok + ' exported' + (failed.length ? ', failed: ' + failed.join(', ') : '.'));
    saveCsv();
    if (results.length) log('Saved ECM_Alert_Info CSV to Downloads. "Copy for Excel" copies the same table.');
    running = false;
  }

  btn('Start', run);
  btn('Stop', function () { stopped = true; });
  btn('Copy for Excel', copyForExcel);
  btn('Load my queue', function () {
    var ids = queueIds();
    box.value = ids.join('\n');
    log(ids.length ? 'Loaded ' + ids.length + ' IDs from My Case Queue.' : 'No queue found - open the Dashboard, or type IDs.');
  });
  btn('Close', function () { stopped = true; panel.remove(); });

  document.body.appendChild(panel);
  var initial = queueIds();
  if (initial.length) { box.value = initial.join('\n'); log('Loaded ' + initial.length + ' IDs from My Case Queue. Remove any you do not want, then press Start.'); }
  box.focus();
})();
