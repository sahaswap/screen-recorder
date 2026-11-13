/*
  ECM Narrative Probe
  Run on one case's Requirements page (or any ECM case page) while logged in. It finds where the
  narrative is kept and lists the page's structure: field labels, element ids, the kind of box each
  field uses, how long its text is, the buttons that generate or download things, and the other
  case pages linked from this one. It never shows what a field says - only how long it is - and
  any run of 5 or more digits is shown as #. Copy the report and send it back to build the
  narrative download into ecm-export.js.
*/
(function () {
  if (document.getElementById('__ecmProbe')) return;

  var VERSION = '2026-11-13';
  /* Words that name a narrative-like field or heading. */
  var WORDS = /narr|summary|comment|justif|disposition|rationale|conclusion|decision|note/i;
  /* Words on buttons and links that make or fetch a document. */
  var ACTIONS = /narr|generat|download|export|print|pdf|word|docx|attach|report/i;
  /* A box a field's text can be in. */
  var CONTROL = 'textarea, input:not([type=hidden]):not([type=button]):not([type=submit]):not([type=checkbox]):not([type=radio]), ' +
    'select, [contenteditable="true"], [contenteditable=""], iframe';
  var out = [];
  var cands = [];

  function line(s) { out.push(s); }
  function mask(s) { return String(s == null ? '' : s).replace(/\d{5,}/g, '#').replace(/\s+/g, ' ').trim(); }
  function short(s, n) { s = mask(s); return s.length > n ? s.slice(0, n) + '...' : s; }
  function shown(el) { return el.getClientRects().length > 0; }
  function addCand(el) { if (el && el.id && cands.indexOf(el) < 0) cands.push(el); }

  function desc(el) {
    var s = el.tagName.toLowerCase();
    if (el.tagName === 'INPUT' && el.type) s += '[type=' + el.type + ']';
    if (el.id) s += ' id="' + mask(el.id) + '"';
    if (el.getAttribute('name')) s += ' name="' + mask(el.getAttribute('name')) + '"';
    if (typeof el.className === 'string' && el.className.trim()) s += ' class="' + short(el.className, 60) + '"';
    if (!shown(el)) s += ' (not shown)';
    return s;
  }

  /* How much text a box holds: -1 when it is a frame from another site, which can't be read. */
  function textLen(el) {
    if (!el) return 0;
    if (el.tagName === 'IFRAME') {
      try {
        var b = el.contentDocument && el.contentDocument.body;
        return b ? (b.innerText || '').length : 0;
      } catch (e) { return -1; }
    }
    if (el.tagName === 'SELECT') return el.selectedIndex >= 0 ? (el.options[el.selectedIndex].text || '').length : 0;
    if ('value' in el && el.tagName !== 'BUTTON' && el.tagName !== 'LI') return (el.value || '').length;
    return (el.innerText || '').length;
  }
  function lenText(el) {
    var n = textLen(el);
    return n < 0 ? 'text not readable (frame from another site)' : 'text ' + n + ' characters';
  }

  /* The box in a field's area, with any hidden input beside it (rich text editors keep their text there). */
  function controlInfo(area) {
    var c = area.querySelector(CONTROL);
    var hid = area.querySelectorAll('input[type=hidden]');
    var s = c ? desc(c) + ', ' + lenText(c) : 'no box, ' + 'text ' + (area.innerText || '').length + ' characters';
    if (c && c.tagName === 'TEXTAREA') s += ', ' + c.rows + ' rows';
    Array.prototype.forEach.call(hid, function (h) {
      if ((h.value || '').length > 0) s += '; hidden ' + desc(h) + ' value ' + h.value.length + ' characters';
    });
    return { text: s, box: c };
  }

  /* The Requirements field label a box sits under, if any. */
  function fieldLabel(el) {
    for (var a = el; a && a !== document.body; a = a.parentElement) {
      var m = /^(.*)divSoureField_(\d+)$/.exec(a.id || '');
      if (m) {
        var lab = document.getElementById(m[1] + 'ReqLabel_' + m[2]);
        return lab ? short(lab.textContent, 60) : '';
      }
    }
    if (el.id) {
      var l = document.querySelector('label[for="' + el.id.replace(/"/g, '') + '"]');
      if (l) return short(l.textContent, 60);
    }
    return '';
  }

  function pathOf(href) {
    try {
      var u = new URL(href, location.href);
      return u.origin === location.origin ? mask(u.pathname + u.search) : '(other site) ' + mask(u.origin + u.pathname);
    } catch (e) { return short(href, 80); }
  }

  /* ---- 1. The page */
  line('ECM Narrative Probe v' + VERSION);
  line('Page: ' + pathOf(location.href) + '   Title: ' + short(document.title, 60));
  line('');

  /* ---- 2. Requirements fields: label -> box */
  var labels = document.querySelectorAll('[id*="ReqLabel_"]');
  line('== Requirement fields (' + labels.length + ') ==');
  Array.prototype.forEach.call(labels, function (lab) {
    var m = /^(.*)ReqLabel_(\d+)$/.exec(lab.id);
    var field = m ? document.getElementById(m[1] + 'divSoureField_' + m[2]) : null;
    var info = field ? controlInfo(field) : null;
    line((m ? m[2] : '?') + '. "' + short(lab.textContent, 70) + '" -> ' +
      (info ? info.text : 'no divSoureField_' + (m ? m[2] : '?')));
    if (info && WORDS.test(lab.textContent)) addCand(info.box || field);
  });
  line('');

  /* ---- 3. Elements whose id, name or class says narrative */
  var named = Array.prototype.filter.call(document.querySelectorAll('[id], [name], [class]'), function (el) {
    return /narr/i.test((el.id || '') + ' ' + (el.getAttribute('name') || '') + ' ' +
      (typeof el.className === 'string' ? el.className : ''));
  }).slice(0, 40);
  line('== Elements named like "narrative" (' + named.length + ') ==');
  named.forEach(function (el) {
    var lab = fieldLabel(el);
    line(desc(el) + ', ' + lenText(el) + (lab ? ', under "' + lab + '"' : ''));
    addCand(el);
  });
  line('');

  /* ---- 4. Short labels and headings with narrative-like words, and the box near each */
  var heads = Array.prototype.filter.call(document.querySelectorAll('body *'), function (el) {
    if (/^(SCRIPT|STYLE|OPTION|SELECT|TEXTAREA|NOSCRIPT)$/.test(el.tagName)) return false;
    var t = (el.textContent || '').trim();
    return t.length > 0 && t.length < 80 && WORDS.test(t);
  });
  heads = heads.filter(function (el) { return !heads.some(function (o) { return o !== el && el.contains(o); }); }).slice(0, 40);
  line('== Labels / headings with narrative-like words (' + heads.length + ') ==');
  heads.forEach(function (el) {
    var near = null;
    for (var a = el.parentElement, k = 0; a && k < 4 && !near; a = a.parentElement, k++) near = a.querySelector(CONTROL);
    line('"' + short(el.textContent, 60) + '" in ' + desc(el) +
      (near ? ' -> nearest box ' + desc(near) + ', ' + lenText(near) : ' -> no box nearby'));
    if (near) addCand(near);
  });
  line('');

  /* ---- 5. Every box that holds a lot of text */
  var boxes = Array.prototype.filter.call(document.querySelectorAll(
    'textarea, [contenteditable="true"], [contenteditable=""], iframe, input[type=hidden]'), function (el) {
    return el.tagName !== 'INPUT' || (el.value || '').length > 200;
  });
  var blocks = Array.prototype.filter.call(document.querySelectorAll('div, td, span, p, pre'), function (el) {
    var n = (el.innerText || '').length;
    if (n < 300) return false;
    return !Array.prototype.some.call(el.children, function (c) { return (c.innerText || '').length > n * 0.8; });
  });
  var big = boxes.concat(blocks).sort(function (a, b) { return textLen(b) - textLen(a); }).slice(0, 25);
  line('== Boxes and blocks holding the most text (' + big.length + ') ==');
  big.forEach(function (el) {
    var lab = fieldLabel(el);
    line(desc(el) + ', ' + lenText(el) + (lab ? ', under "' + lab + '"' : ''));
    if (el.tagName !== 'INPUT' && textLen(el) >= 300) addCand(el);
  });
  line('');

  /* ---- 6. Buttons and links that make or fetch a document */
  var acts = Array.prototype.filter.call(document.querySelectorAll(
    'a, button, input[type=button], input[type=submit], input[type=image], [onclick]'), function (el) {
    var t = (el.innerText || el.value || '') + ' ' + (el.id || '') + ' ' + (el.title || '') + ' ' + (el.getAttribute('alt') || '');
    return ACTIONS.test(t);
  }).slice(0, 40);
  line('== Buttons / links that generate, print or download (' + acts.length + ') ==');
  acts.forEach(function (el) {
    var pb = /__doPostBack\('([^']*)'/.exec((el.getAttribute('onclick') || '') + ' ' + (el.getAttribute('href') || ''));
    var href = el.getAttribute('href');
    line('"' + short(el.innerText || el.value || el.title || '', 50) + '" ' + desc(el) +
      (pb ? ', posts back ' + mask(pb[1]) : '') +
      (href && !/^javascript:/i.test(href) && href !== '#' ? ', goes to ' + pathOf(href) : '') +
      (el.getAttribute('onclick') && !pb ? ', onclick: ' + short(el.getAttribute('onclick'), 80) : ''));
  });
  line('');

  /* ---- 7. Tabs, and the other case pages linked from here */
  var tabs = document.querySelectorAll('[role=tab], .nav-tabs a, .rtsLink, .tab a, .tabs a');
  line('== Tabs (' + tabs.length + ') ==');
  line(Array.prototype.map.call(tabs, function (t) { return '"' + short(t.innerText || t.textContent, 40) + '"'; }).join(', '));
  var pages = [];
  Array.prototype.forEach.call(document.querySelectorAll('a[href*="caseid="], a[href*="CaseId="], a[href*=".aspx"]'), function (a) {
    var p = pathOf(a.getAttribute('href')).replace(/\?.*$/, '');
    if (pages.indexOf(p) < 0) pages.push(p);
  });
  line('== Pages linked from here (' + pages.length + ') ==');
  line(pages.slice(0, 60).join('\n'));
  line('');

  /* ---- 8. Frames */
  var frames = document.querySelectorAll('iframe, frame');
  line('== Frames (' + frames.length + ') ==');
  Array.prototype.forEach.call(frames, function (f) {
    line(desc(f) + ', src ' + pathOf(f.getAttribute('src') || 'about:blank') + ', ' + lenText(f));
  });
  line('');

  var panel = document.createElement('div');
  panel.id = '__ecmProbe';
  panel.style.cssText = 'position:fixed;right:16px;bottom:16px;z-index:2147483647;width:520px;box-sizing:border-box;' +
    'background:#1f2937;color:#f9fafb;padding:12px;border-radius:10px;box-shadow:0 6px 24px rgba(0,0,0,.4);' +
    'font:13px/1.4 Segoe UI,Arial,sans-serif;text-align:left';
  panel.innerHTML =
    '<div style="display:flex;justify-content:space-between;align-items:center;margin:0 0 8px">' +
    '<b style="font-size:14px">ECM Narrative Probe</b>' +
    '<span class="__probeNote" style="font-size:12px;color:#9ca3af">checking the raw page...</span></div>' +
    '<textarea readonly rows="18" style="display:block;width:100%;box-sizing:border-box;margin:0;padding:6px 8px;' +
    'border:1px solid #4b5563;border-radius:6px;background:#111827;color:#e5e7eb;font:12px/1.45 Consolas,monospace;' +
    'white-space:pre;resize:vertical"></textarea>' +
    '<div style="display:grid;grid-template-columns:1fr 1fr;gap:6px;margin:8px 0 0">' +
    '<button class="__probeCopy" style="all:unset;text-align:center;height:30px;line-height:30px;border-radius:6px;' +
    'cursor:pointer;background:#2563eb;color:#fff;font-weight:600">Copy report</button>' +
    '<button class="__probeClose" style="all:unset;text-align:center;height:30px;line-height:30px;border-radius:6px;' +
    'cursor:pointer;background:#e5e7eb;color:#111827">Close</button></div>' +
    '<div style="margin:8px 0 0;font-size:12px;color:#9ca3af">Shows field names and lengths only, never what a ' +
    'field says. Read it through before sending it.</div>';
  var ta = panel.querySelector('textarea');
  var note = panel.querySelector('.__probeNote');
  function show() { ta.value = out.join('\n'); }
  panel.querySelector('.__probeClose').onclick = function (e) { e.preventDefault(); panel.remove(); };
  panel.querySelector('.__probeCopy').onclick = function (e) {
    e.preventDefault();
    var done = function () { note.textContent = 'Copied.'; };
    var fallback = function () { ta.select(); document.execCommand('copy'); done(); };
    if (navigator.clipboard) navigator.clipboard.writeText(ta.value).then(done, fallback); else fallback();
  };
  document.body.appendChild(panel);
  show();

  /* ---- 9. Are the candidates in the page as the server sends it? The export reads pages that way
     (fetch, no scripts run), so a narrative only filled in by the page's scripts needs a slower read. */
  line('== In the raw page from the server (' + cands.length + ' candidate(s)) ==');
  fetch(location.href, { credentials: 'same-origin' }).then(function (res) {
    if (/\/Login/i.test(res.url)) throw new Error('sent to the login page');
    return res.text();
  }).then(function (html) {
    var d = new DOMParser().parseFromString(html, 'text/html');
    line('Requirement labels in raw page: ' + d.querySelectorAll('[id*="ReqLabel_"]').length + ' (live page: ' + labels.length + ')');
    cands.forEach(function (el) {
      var r = d.getElementById(el.id);
      var n = 0;
      if (r) {
        if (r.tagName === 'TEXTAREA') n = r.textContent.length;
        else if (r.tagName === 'INPUT') n = (r.getAttribute('value') || '').length;
        else n = r.textContent.length;
      }
      line(desc(el) + ': ' + (r ? 'there, raw text ' + n + ' characters (live ' + textLen(el) + ')' : 'NOT in the raw page'));
    });
    note.textContent = 'Done.';
    show();
  }).catch(function (e) {
    line('Could not fetch the raw page: ' + e.message);
    note.textContent = 'Done (raw page not checked).';
    show();
  });
})();
