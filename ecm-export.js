/*
  ECM Transactions Export
  For each alert ID, opens the alert's Statistics page and runs "Export Transactions",
  saving the file to Downloads as <alertId>_<original file name>.
  Run on any ECM page while logged in: paste into the Console (Ctrl+Shift+J) or use the bookmarklet.
  Uses your existing Edge session - no passwords are stored or sent anywhere else.
*/
(function () {
  if (document.getElementById('__ecmExport')) return;

  var BASE = location.origin + '/ECM/Case/Statistics.aspx?caseid=';
  var BUTTON_ID = 'btnExportStatsTransactions';
  var TAB_WAIT_MS = 10000; /* fallback mode: time allowed for a download to start before closing the tab */
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

  /* Main mode: load the page and post the export in the background (no tabs). */
  async function viaFetch(id) {
    var url = BASE + id;
    var res = await fetch(url, { credentials: 'same-origin' });
    if (isLogin(res.url)) throw new Error('LOGIN');
    if (!res.ok) return 'page error ' + res.status;
    var doc = new DOMParser().parseFromString(await res.text(), 'text/html');
    var btn = doc.getElementById(BUTTON_ID);
    if (!btn) return 'no Export button on page';
    var form = btn.form || btn.closest('form');
    if (!form) return 'no form';
    var action = new URL(form.getAttribute('action') || res.url, res.url).href;
    var post = await fetch(action, { method: 'POST', credentials: 'same-origin', body: formBody(form, btn) });
    if (isLogin(post.url)) throw new Error('LOGIN');
    var type = post.headers.get('content-type') || '';
    var disp = post.headers.get('content-disposition') || '';
    if (post.ok && (/attachment/i.test(disp) || !/text\/html/i.test(type))) {
      var blob = await post.blob();
      /* An error page sent without a proper content type is not a file. */
      if (!/attachment/i.test(disp) && /^\s*<(!doctype|html)/i.test(await blob.slice(0, 200).text())) return 'server did not return a file';
      var name = fileName(disp, type, id);
      saveBlob(blob, name);
      return 'OK ' + name;
    }
    return 'server did not return a file';
  }

  /* Fallback mode: open the page in a tab and click the button, like a person would. */
  async function viaTab(id) {
    var w = window.open(BASE + id, '_blank');
    if (!w) return 'pop-up blocked - allow pop-ups for this site';
    try {
      var btn = null;
      for (var i = 0; i < 120 && !btn; i++) {
        await sleep(500);
        try { if (w.document.readyState === 'complete') btn = w.document.getElementById(BUTTON_ID); } catch (e) { /* still loading */ }
      }
      if (!btn) return 'no Export button on page';
      btn.click();
      await sleep(TAB_WAIT_MS);
      return 'OK (tab mode, check Downloads)';
    } finally {
      try { w.close(); } catch (e) { /* already closed */ }
    }
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
    var ok = 0, failed = [];
    for (var i = 0; i < ids.length && !stopped; i++) {
      var id = ids[i];
      var line = log('[' + (i + 1) + '/' + ids.length + '] ' + id + ' ...');
      var result;
      try {
        result = await viaFetch(id);
        if (result.indexOf('OK') !== 0) {
          line.textContent += ' ' + result + ', trying tab mode ...';
          result = await viaTab(id);
        }
      } catch (e) {
        if (e.message === 'LOGIN') { line.textContent += ' session expired'; log('Log in to ECM again, then press Start.'); break; }
        result = 'error: ' + e.message;
      }
      line.textContent = '[' + (i + 1) + '/' + ids.length + '] ' + id + ' ' + result;
      if (result.indexOf('OK') === 0) ok++; else failed.push(id);
      await sleep(PAUSE_MS);
    }
    log((stopped ? 'Stopped. ' : 'Finished. ') + ok + ' exported' + (failed.length ? ', failed: ' + failed.join(', ') : '.'));
    running = false;
  }

  btn('Start', run);
  btn('Stop', function () { stopped = true; });
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
