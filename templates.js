/* Sjablonen: tab in het paneel Handtekening.
 * Leest templates/index.json, toont invulvelden voor {{Veld}} en zet het sjabloon op de cursor in de mail.
 */
var NVTpl = (function () {
  var list = [], current = null, body = "", values = {};
  var RE = /\{\{\s*([^{}:]+?)\s*(?::\s*(lang))?\s*\}\}/g;

  function $(id) { return document.getElementById(id); }
  function esc(s) { return String(s).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); }
  function say(t) { $("tplStatus").textContent = t; }

  function fieldsOf(text) {
    var out = [], seen = {}, m;
    RE.lastIndex = 0;
    while ((m = RE.exec(text))) {
      if (!seen[m[1]]) { seen[m[1]] = 1; out.push({ name: m[1], long: !!m[2] }); }
      else if (m[2]) { for (var i = 0; i < out.length; i++) if (out[i].name === m[1]) out[i].long = true; }
    }
    return out;
  }
  function fill(text, html) {
    return text.replace(RE, function (all, name) {
      var v = values[name];
      if (!v || !String(v).trim()) return "[" + name + "]";
      return html ? esc(v).replace(/\r?\n/g, "<br>") : String(v);
    });
  }
  function filledHtml() { return fill(body, true); }
  function filledSubject() { return current && current.subject ? fill(current.subject, false) : ""; }

  function renderList() {
    var box = $("tplList"); box.innerHTML = "";
    if (!list.length) { box.innerHTML = '<p class="hint">Geen sjablonen gevonden.</p>'; return; }
    list.forEach(function (t) {
      var b = document.createElement("button");
      b.type = "button"; b.className = "tpl" + (current && current.id === t.id ? " on" : "");
      b.setAttribute("aria-pressed", String(!!(current && current.id === t.id)));
      var n = document.createElement("b"); n.textContent = t.title; b.appendChild(n);
      if (t.desc) { var d = document.createElement("span"); d.textContent = t.desc; b.appendChild(d); }
      b.onclick = function () { open(t); };
      box.appendChild(b);
    });
  }
  function renderFields() {
    var box = $("tplFields"); box.innerHTML = "";
    var f = fieldsOf((current.subject || "") + "\n" + body);
    $("tplFieldsCard").hidden = !f.length;
    f.forEach(function (x) {
      var l = document.createElement("label"), el = document.createElement(x.long ? "textarea" : "input");
      if (x.long) el.rows = 4; else el.type = "text";
      el.value = values[x.name] || "";
      el.oninput = function () { values[x.name] = el.value; renderPreview(); };
      l.appendChild(document.createTextNode(x.name)); l.appendChild(el); box.appendChild(l);
    });
  }
  function renderPreview() {
    $("tplSubject").textContent = filledSubject() || "(geen onderwerp)";
    $("tplPreview").innerHTML = filledHtml();
  }
  function open(t) {
    current = t; values = {}; say("Laden...");
    fetch("templates/" + encodeURIComponent(t.file) + "?t=" + Date.now(), { cache: "no-store" })
      .then(function (r) { if (!r.ok) throw new Error(r.status); return r.text(); })
      .then(function (txt) {
        body = txt; renderList(); renderFields(); renderPreview();
        $("tplDetail").hidden = false; say("");
        var first = $("tplFields").querySelector("input,textarea"); if (first) first.focus();
      })
      .catch(function (e) { $("tplDetail").hidden = true; say("Sjabloon laden lukte niet (" + e.message + ")."); });
  }

  function insert() {
    if (!current) return;
    var item = window.Office && Office.context && Office.context.mailbox && Office.context.mailbox.item;
    if (!item || !item.body || !item.body.setSelectedDataAsync) { say("Invoegen werkt alleen in een mail die je aan het schrijven bent. Gebruik anders Kopieer."); return; }
    var html = '<div>' + filledHtml() + '</div>';
    item.body.setSelectedDataAsync(html, { coercionType: Office.CoercionType.Html }, function (r) {
      if (r.status !== Office.AsyncResultStatus.Succeeded) { say("Invoegen lukte niet: " + (r.error && r.error.message || "onbekende fout")); return; }
      var subj = filledSubject();
      if (!subj || !$("tplSetSubject").checked || !item.subject || !item.subject.setAsync) { say("Sjabloon ingevoegd."); return; }
      item.subject.setAsync(subj, function (s) {
        say(s.status === Office.AsyncResultStatus.Succeeded ? "Sjabloon en onderwerp ingevoegd." : "Sjabloon ingevoegd, onderwerp lukte niet.");
      });
    });
  }
  function copy() {
    if (!current) return;
    var html = filledHtml(), box = $("tplPreview"), text = (box.innerText || box.textContent || "").trim();
    var done = function () { say("Sjabloon gekopieerd. Plak het waar je het nodig hebt."); };
    var fallback = function () {
      var ok = false, sel = window.getSelection(), r = document.createRange();
      try { r.selectNodeContents(box); sel.removeAllRanges(); sel.addRange(r); ok = document.execCommand("copy"); } catch (e) {}
      if (ok) { sel.removeAllRanges(); done(); } else say("Tekst is geselecteerd. Druk op Ctrl+C om te kopiëren.");
    };
    if (window.ClipboardItem && navigator.clipboard && navigator.clipboard.write) {
      navigator.clipboard.write([new ClipboardItem({ "text/html": new Blob([html], { type: "text/html" }), "text/plain": new Blob([text], { type: "text/plain" }) })]).then(done, fallback);
    } else fallback();
  }

  function init() {
    $("tplInsert").onclick = insert;
    $("tplCopy").onclick = copy;
    fetch("templates/index.json?t=" + Date.now(), { cache: "no-store" })
      .then(function (r) { if (!r.ok) throw new Error(r.status); return r.json(); })
      .then(function (j) { list = Array.isArray(j) ? j : []; renderList(); })
      .catch(function (e) { list = []; renderList(); say("Lijst met sjablonen laden lukte niet (" + e.message + ")."); });
  }

  return { init: init, _fill: fill, _fields: fieldsOf, _set: function (t, b, v) { current = t; body = b; values = v || {}; } };
})();
