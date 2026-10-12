/*
  VDI Download Recorder
  Records the steps you take on a website (pages, clicks, form fields, downloads)
  so a download macro can be built from them.
  Typed values are NEVER recorded - only which field was used.
  Run: paste into the browser console (F12 > Console), or use the bookmarklet.
  Run again (or click Stop) to stop. Log survives page reloads and is shared by all tabs on the same site.
*/
(function () {
  if (window.__vdiRec) { window.__vdiRec.stop(); return; }

  var KEY = '__vdiRecLog';
  var FILE_RX = /\.(xlsx?|xlsm|csv|pdf|zip|txt|docx?|xml|json|dat)(\?|#|$)/i;
  var log = load();
  var lastUrl = location.href;

  function load() { try { return JSON.parse(localStorage.getItem(KEY)) || []; } catch (e) { return []; } }
  function save() { try { localStorage.setItem(KEY, JSON.stringify(log)); } catch (e) {} }
  /* Re-read before every write so several tabs on the same site share one log. */
  function add(ev) { ev.at = new Date().toISOString(); log = load(); log.push(ev); save(); render(); }

  function clean(s) { return (s || '').trim().replace(/\s+/g, ' ').slice(0, 80); }

  /* Visible label of an element. Never reads what the user typed. */
  function text(el) {
    if (!el || !el.tagName) return '';
    var t = el.tagName;
    if (t === 'INPUT' && /^(button|submit|reset)$/i.test(el.type)) return clean(el.value);
    if (t === 'INPUT' || t === 'TEXTAREA' || t === 'SELECT') {
      return clean(el.getAttribute('aria-label') || el.placeholder || el.name || el.id);
    }
    return clean(el.innerText || el.getAttribute('aria-label') || el.title);
  }

  function labelFor(el) {
    if (el.labels && el.labels[0]) return clean(el.labels[0].innerText);
    return clean(el.getAttribute('aria-label') || el.placeholder);
  }

  /* Short CSS selector a macro can use to find the element again. */
  function sel(el) {
    var parts = [];
    while (el && el.nodeType === 1 && parts.length < 5) {
      if (el.id) { parts.unshift('#' + CSS.escape(el.id)); break; }
      var p = el.tagName.toLowerCase();
      var name = el.getAttribute('name');
      if (name) {
        p += '[name="' + name + '"]';
      } else {
        var i = 1, s = el;
        while ((s = s.previousElementSibling)) if (s.tagName === el.tagName) i++;
        p += ':nth-of-type(' + i + ')';
      }
      parts.unshift(p);
      el = el.parentElement;
    }
    return parts.join(' > ');
  }

  function fileLinks() {
    return Array.prototype.slice.call(document.querySelectorAll('a[href]'))
      .filter(function (a) { return a.hasAttribute('download') || FILE_RX.test(a.href); })
      .slice(0, 50)
      .map(function (a) { return { text: text(a), href: a.href }; });
  }

  function pageEvent() {
    add({
      type: 'page',
      url: location.href,
      title: document.title,
      iframes: document.querySelectorAll('iframe').length,
      fileLinks: fileLinks()
    });
  }

  function onClick(e) {
    if (panel.contains(e.target)) return;
    var el = e.target.closest('a,button,input,select,textarea,label,[role=button],[role=link],[role=tab],[role=menuitem],[onclick]') || e.target;
    var ev = { type: 'click', tag: el.tagName.toLowerCase(), text: text(el), selector: sel(el) };
    if (el.href) ev.href = el.href;
    if (el.hasAttribute('download')) ev.download = el.getAttribute('download') || true;
    if (el.type) ev.inputType = el.type;
    add(ev);
  }

  function onChange(e) {
    var el = e.target;
    if (panel.contains(el) || !/^(INPUT|SELECT|TEXTAREA)$/.test(el.tagName)) return;
    var ev = {
      type: 'field', tag: el.tagName.toLowerCase(), inputType: el.type || '',
      name: el.name || '', id: el.id || '', label: labelFor(el), selector: sel(el)
    };
    if (el.tagName === 'SELECT') ev.chose = el.options[el.selectedIndex] ? clean(el.options[el.selectedIndex].text) : '';
    else if (el.type === 'checkbox' || el.type === 'radio') ev.checked = el.checked;
    else if (el.type === 'date') ev.value = el.value; /* dates help spot "pick report date" steps */
    else ev.value = '<typed ' + (el.value || '').length + ' chars - not recorded>';
    add(ev);
  }

  function onSubmit(e) {
    var f = e.target;
    add({ type: 'submit', action: f.action || '', method: (f.method || 'get').toUpperCase(), selector: sel(f) });
  }

  var timer = setInterval(function () {
    if (location.href !== lastUrl) { lastUrl = location.href; pageEvent(); }
    else { var n = load().length; if (n !== log.length) { log = load(); render(); } }
  }, 500);

  /* ---------- floating control panel ---------- */
  var panel = document.createElement('div');
  panel.style.cssText = 'position:fixed;right:12px;bottom:12px;z-index:2147483647;background:#1f2937;color:#fff;' +
    'font:13px/1.4 Segoe UI,Arial,sans-serif;padding:10px 12px;border-radius:8px;box-shadow:0 4px 16px rgba(0,0,0,.35)';
  var status = document.createElement('div');
  panel.appendChild(status);
  var row = document.createElement('div');
  row.style.marginTop = '6px';
  panel.appendChild(row);

  function btn(label, fn) {
    var b = document.createElement('button');
    b.textContent = label;
    b.style.cssText = 'margin-right:4px;padding:3px 8px;border:0;border-radius:4px;cursor:pointer;background:#e5e7eb;color:#111';
    b.onclick = function (e) { e.preventDefault(); e.stopPropagation(); fn(); };
    row.appendChild(b);
  }

  function output() {
    log = load();
    return JSON.stringify({ recorder: 'vdi-rec v1', browser: navigator.userAgent, events: log }, null, 2);
  }

  function copy() {
    var s = output();
    function fallback() {
      var ta = document.createElement('textarea');
      ta.value = s; document.body.appendChild(ta); ta.select();
      document.execCommand('copy'); ta.remove();
      alert('Recording copied to clipboard (' + log.length + ' events).');
    }
    if (navigator.clipboard) navigator.clipboard.writeText(s).then(function () {
      alert('Recording copied to clipboard (' + log.length + ' events).');
    }, fallback);
    else fallback();
  }

  function saveFile() {
    var a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([output()], { type: 'application/json' }));
    a.download = 'recording.json';
    document.body.appendChild(a); a.click(); a.remove();
  }

  function note() {
    var n = prompt('Add a note (e.g. "waited 30s for report to generate", "logged in via SSO"):');
    if (n) add({ type: 'note', text: n });
  }

  function clearLog() {
    if (confirm('Clear all recorded steps?')) { log = []; save(); pageEvent(); }
  }

  function render() { status.textContent = '● REC  ' + log.length + ' steps'; status.style.color = '#f87171'; }

  function stop() {
    clearInterval(timer);
    document.removeEventListener('click', onClick, true);
    document.removeEventListener('change', onChange, true);
    document.removeEventListener('submit', onSubmit, true);
    panel.remove();
    delete window.__vdiRec;
  }

  btn('Note', note);
  btn('Copy', copy);
  btn('Save', saveFile);
  btn('Clear', clearLog);
  btn('Stop', stop);

  document.body.appendChild(panel);
  document.addEventListener('click', onClick, true);
  document.addEventListener('change', onChange, true);
  document.addEventListener('submit', onSubmit, true);
  window.__vdiRec = { stop: stop };
  pageEvent();
})();
