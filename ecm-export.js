/*
  ECM Transactions Export
  For each alert ID (ECM case ID):
    1. reads Customer Name, Customer Internal ID, current and permanent address, and Velocity Alert ID from the Requirements page;
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
  var COLUMNS = ['ECM ID', 'Velocity Alert ID', 'Customer Name', 'Customer Internal ID', 'Address',
    'Street', 'City', 'State', 'Zip Code', 'Country', 'Permanent Address', 'Transaction File', 'Result'];
  var results = [];
  var BUTTON_ID = 'btnExportStatsTransactions';
  var CLICK_WAIT_MS = 10000; /* after clicking Export: time allowed for the download to start */
  var PAUSE_MS = 300;      /* pause between alerts, to go easy on the server */
  var POLL_MS = 200;       /* how often a loading page is checked */
  var PAGE_TIMEOUT_MS = 60000;
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

  function readInfo(id) {
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
        (r.permanent || {}).full, r.file, r.result];
    });
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
  /* Every style is set inline and in full so the ECM page's own CSS cannot shift the layout. */
  var FONT = 'font-family:Segoe UI,Arial,sans-serif;font-size:13px;line-height:1.4;letter-spacing:normal;text-transform:none;';
  panel.style.cssText = 'position:fixed;right:16px;bottom:16px;z-index:2147483647;width:360px;box-sizing:border-box;' +
    'background:#1f2937;color:#f9fafb;padding:12px;border-radius:10px;box-shadow:0 6px 24px rgba(0,0,0,.4);text-align:left;' + FONT;
  panel.innerHTML =
    '<div style="display:flex;align-items:center;justify-content:space-between;margin:0 0 8px;' + FONT + '">' +
      '<span style="font-weight:600;font-size:14px">ECM Export Transactions</span>' +
      '<span style="display:flex;align-items:center;gap:8px">' +
        '<span class="count" style="font-size:12px;color:#9ca3af"></span>' +
        '<button class="x" title="Close" style="all:unset;cursor:pointer;width:22px;height:22px;line-height:22px;text-align:center;' +
          'border-radius:4px;color:#d1d5db;font:16px Arial,sans-serif">&#x2715;</button>' +
      '</span>' +
    '</div>' +
    '<textarea rows="6" placeholder="Alert IDs - one per line, or separated by commas/spaces" style="display:block;width:100%;' +
      'box-sizing:border-box;margin:0;padding:6px 8px;border:1px solid #4b5563;border-radius:6px;background:#f9fafb;color:#111827;' +
      'font:13px/1.4 Consolas,monospace;resize:vertical"></textarea>' +
    '<div class="row" style="display:grid;grid-template-columns:1fr 1fr;gap:6px;margin:8px 0"></div>' +
    '<div class="log" style="max-height:180px;overflow:auto;padding:6px 8px;border-radius:6px;background:#111827;color:#e5e7eb;' +
      'font:12px/1.45 Consolas,monospace;white-space:pre-wrap;min-height:18px"></div>';
  var box = panel.querySelector('textarea');
  var row = panel.querySelector('.row');
  var logEl = panel.querySelector('.log');
  var countEl = panel.querySelector('.count');

  function updateCount() {
    var n = parseIds(box.value).length;
    countEl.textContent = n ? n + (n === 1 ? ' ID' : ' IDs') : '';
  }
  box.addEventListener('input', updateCount);

  function btn(label, fn, primary) {
    var b = document.createElement('button');
    b.textContent = label;
    b.style.cssText = 'all:unset;box-sizing:border-box;display:block;width:100%;height:30px;line-height:30px;text-align:center;' +
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
      var line = log(tag + ' working ...');
      var info = {}, result, started = Date.now();
      try {
        /* Details and export pages load side by side. */
        var both = await Promise.all([readInfo(id), getExport(id)]);
        info = both[0];
        var ex = both[1];
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
        if (e.message === 'LOGIN') { line.textContent = tag + ' session expired'; log('Log in to ECM again, then press Start.'); break; }
        result = 'error: ' + e.message;
      }
      line.textContent = tag + ' ' + result + '  (' + ((Date.now() - started) / 1000).toFixed(1) + ' s)';
      results.push({
        ecm: id, velocity: info.velocity || '', customer: info.customer || '', internalId: info.internalId || '',
        address: info.address, permanent: info.permanent,
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

  btn('Start', run, true);
  btn('Stop', function () { stopped = true; });
  btn('Load my queue', function () {
    var ids = queueIds();
    box.value = ids.join('\n');
    updateCount();
    log(ids.length ? 'Loaded ' + ids.length + ' IDs from My Case Queue.' : 'No queue found - open the Dashboard, or type IDs.');
  });
  btn('Copy for Excel', copyForExcel);
  panel.querySelector('.x').onclick = function (e) { e.preventDefault(); stopped = true; panel.remove(); };

  document.body.appendChild(panel);
  var initial = queueIds();
  if (initial.length) { box.value = initial.join('\n'); updateCount(); log('Loaded ' + initial.length + ' IDs from My Case Queue. Remove any you do not want, then press Start.'); }
  box.focus();
})();
