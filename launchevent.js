/* Nico Vermaat e-mailhandtekening: gedeelde logica.
 * ES5 zonder DOM, zodat dit ook draait in de JavaScript runtime van klassiek Outlook.
 * https://nvermaat.github.io/vermaat-handtekening wordt bij installatie vervangen door het adres waar de invoegtoepassing staat.
 */
var NVSig = (function () {
  var HOST = "https://nvermaat.github.io/vermaat-handtekening";
  var KEY = "nvsig";

  var DEFAULTS = {
    version: 1,
    autoInsert: true,          // handtekening invoegen bij nieuwe mail
    disableOutlookSig: true,   // eigen Outlook handtekening uitzetten, voorkomt dubbele handtekening
    defaultVariant: "auto",    // auto | prive | consultancy | it
    fields: {
      name: "Nico Vermaat",
      role: "Senior IT Consultant",
      slogan: "Vermaat IT: inzichtelijk, stabiel, opgeleverd.",
      bizphone: "+31 6 5366 2022",
      phone: "+31 6 2705 1514",
      emailBiz: "info@vermaat-it.nl",
      emailPriv: "info@nicovermaat.nl",
      address: "Utrechtlaan 11\n3137 RG Vlaardingen",
      web: "www.vermaat-it.nl",
      kvk: "73918024",
      company: "Nico Vermaat IT Consultancy",
      company2: "Nico Vermaat IT",
      activity: "IT-consultancy",
      certs: "MCSE · VCP DCV · VCP DTM · CCE"
    },
    show: {
      prive:       { logo: false, role: false, slogan: false, bizphone: false, privphone: true,  addr: true,  web: false, certs: false, kvk: false, footer: false },
      consultancy: { logo: true,  role: true,  slogan: true,  bizphone: true,  privphone: false, addr: false, web: true,  certs: true,  kvk: true,  footer: true },
      it:          { logo: true,  role: true,  slogan: true,  bizphone: true,  privphone: false, addr: false, web: true,  certs: true,  kvk: true,  footer: true }
    }
  };

  var VARIANTS = [
    { id: "prive", label: "Privé" },
    { id: "consultancy", label: "Nico Vermaat IT Consultancy" },
    { id: "it", label: "Nico Vermaat IT" }
  ];

  function clone(o) { return JSON.parse(JSON.stringify(o)); }

  function merge(base, extra) {
    var out = clone(base);
    if (!extra || typeof extra !== "object") return out;
    for (var k in extra) {
      if (!Object.prototype.hasOwnProperty.call(extra, k)) continue;
      if (out[k] && typeof out[k] === "object" && !Array.isArray(out[k])) out[k] = merge(out[k], extra[k]);
      else out[k] = extra[k];
    }
    return out;
  }

  function load(roaming) {
    var saved = null;
    try { saved = roaming && roaming.get(KEY); if (typeof saved === "string") saved = JSON.parse(saved); } catch (e) { saved = null; }
    return merge(DEFAULTS, saved);
  }

  function save(roaming, settings, cb) {
    roaming.set(KEY, JSON.stringify(settings));
    roaming.saveAsync(function (r) { if (cb) cb(r); });
  }

  // Kies de versie op basis van het Van-adres, tenzij er een vaste standaard is gekozen.
  function pickVariant(settings, fromAddress) {
    if (settings.defaultVariant && settings.defaultVariant !== "auto") return settings.defaultVariant;
    var a = String(fromAddress || "").toLowerCase();
    if (/@nicovermaat\.nl$/.test(a)) return "prive";
    if (/@vermaat-it\.nl$/.test(a)) return "consultancy";
    return "consultancy";
  }

  function esc(s) {
    return String(s == null ? "" : s).replace(/[&<>"]/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c];
    });
  }
  function telHref(n) { return "tel:" + String(n).replace(/[^+\d]/g, ""); }
  function trim(s) { return String(s == null ? "" : s).replace(/^\s+|\s+$/g, ""); }

  function build(variant, settings) {
    var f = settings.fields, s = settings.show[variant] || settings.show.consultancy;
    var biz = variant !== "prive";
    var F = "font-family:Arial,Helvetica,sans-serif;";
    var ic = function (file, alt) {
      return '<img src="' + HOST + "/assets/" + file + '" width="12" height="12" alt="' + alt +
        '" style="display:inline-block;width:12px;height:12px;border:0;vertical-align:-1px;">&nbsp;&nbsp;';
    };
    var email = trim(biz ? f.emailBiz : f.emailPriv);
    var web = s.web ? trim(f.web) : "";
    var href = web ? (/^https?:/i.test(web) ? web : "https://" + web.replace(/^\/+/, "")) : "";
    var kvk = s.kvk ? trim(f.kvk) : "";
    var foot = biz && s.footer;

    var phones = [];
    if (s.bizphone && trim(f.bizphone)) phones.push([trim(f.bizphone), "zakelijk"]);
    if (s.privphone && trim(f.phone)) phones.push([trim(f.phone), "privé"]);
    var lbl = phones.length > 1;

    var h = '<table cellpadding="0" cellspacing="0" border="0" role="presentation" style="border-collapse:collapse;' + F + 'color:#0c0c0c;margin-top:14px;">';
    if (s.logo) h += '<tr><td style="padding:0 0 14px 0;"><img src="' + HOST + '/assets/logo.png" width="200" height="34" alt="Nico Vermaat IT Consultancy" style="display:block;border:0;width:200px;height:34px;"></td></tr>';
    h += '<tr><td style="border-left:3px solid #006298;padding:2px 0 2px 12px;">';
    h += '<div style="' + F + 'font-size:15px;font-weight:bold;color:#0c0c0c;line-height:20px;">' + esc(trim(f.name)) + "</div>";
    if (s.role && trim(f.role)) h += '<div style="' + F + 'font-size:13px;color:#55534c;line-height:18px;">' + esc(trim(f.role)) + "</div>";
    if (s.slogan && trim(f.slogan)) h += '<div style="' + F + 'font-size:12px;font-style:italic;color:#006298;line-height:17px;padding-top:3px;">' + esc(trim(f.slogan)) + "</div>";
    h += "</td></tr>";

    for (var i = 0; i < phones.length; i++) {
      h += '<tr><td style="padding:' + (i ? "2px" : "10px") + " 0 0 0;" + F + 'font-size:13px;line-height:18px;">' + ic("phone.png", "Telefoon") +
        '<a href="' + esc(telHref(phones[i][0])) + '" style="color:#0c0c0c;text-decoration:none;">' + esc(phones[i][0]) + "</a>" +
        (lbl ? '<span style="' + F + 'font-size:11px;color:#85826f;">&nbsp;&nbsp;' + phones[i][1] + "</span>" : "") + "</td></tr>";
    }

    var contact = [];
    if (email) contact.push(ic("mail.png", "E-mail") + '<a href="mailto:' + esc(email) + '" style="color:#006298;text-decoration:none;">' + esc(email) + "</a>");
    if (web) contact.push('<a href="' + esc(href) + '" style="color:#006298;text-decoration:none;">' + esc(web) + "</a>");
    if (contact.length) h += '<tr><td style="padding:' + (phones.length ? "2px" : "10px") + " 0 0 0;" + F + 'font-size:13px;line-height:18px;">' +
      contact.join('<span style="color:#006298;font-weight:bold;">&nbsp;/&nbsp;</span>') + "</td></tr>";

    if (s.addr && trim(f.address)) {
      var lines = String(f.address).split(/\n+/), parts = [];
      for (var j = 0; j < lines.length; j++) if (trim(lines[j])) parts.push(esc(trim(lines[j])));
      h += '<tr><td style="padding:6px 0 0 0;' + F + 'font-size:12px;color:#55534c;line-height:17px;">' + parts.join("<br>") + "</td></tr>";
    }
    if (s.certs && trim(f.certs)) h += '<tr><td style="padding:6px 0 0 0;font-family:Consolas,\'Courier New\',monospace;font-size:11px;letter-spacing:0.5px;color:#85826f;line-height:16px;">' + esc(trim(f.certs)) + "</td></tr>";
    if (!foot && kvk) h += '<tr><td style="padding:4px 0 0 0;' + F + 'font-size:11px;color:#85826f;line-height:16px;">KVK ' + esc(kvk) + "</td></tr>";
    if (foot) {
      var co = trim(variant === "it" ? f.company2 : f.company);
      var meta = [];
      if (trim(f.activity)) meta.push(esc(trim(f.activity)));
      if (kvk) meta.push("KVK " + esc(kvk));
      h += '<tr><td style="padding:12px 0 0 0;"><table cellpadding="0" cellspacing="0" border="0" role="presentation" style="border-collapse:collapse;border-top:1px solid #e4e1d6;"><tr><td style="padding:10px 0 0 0;">';
      if (co) h += '<div style="' + F + 'font-size:12px;font-weight:bold;color:#0c0c0c;line-height:17px;">' + esc(co) + "</div>";
      if (meta.length) h += '<div style="' + F + 'font-size:11px;color:#71705f;line-height:16px;">' + meta.join("&nbsp;&nbsp;·&nbsp;&nbsp;") + "</div>";
      h += "</td></tr></table></td></tr>";
    }
    h += "</table>";
    return h;
  }

  return { HOST: HOST, DEFAULTS: DEFAULTS, VARIANTS: VARIANTS, load: load, save: save, merge: merge, pickVariant: pickVariant, build: build };
})();

if (typeof module !== "undefined" && module.exports) module.exports = NVSig;
/* Automatisch invoegen: draait bij een nieuwe mail en wanneer je het Van-adres wijzigt. */
function nvInsertSignature(event) {
  var done = function () { try { event.completed(); } catch (e) {} };
  try {
    var mailbox = Office.context.mailbox;
    var item = mailbox.item;
    var settings = NVSig.load(Office.context.roamingSettings);
    if (!settings.autoInsert) { done(); return; }

    var apply = function (fromAddress) {
      var html = NVSig.build(NVSig.pickVariant(settings, fromAddress), settings);
      var setSig = function () {
        item.body.setSignatureAsync(html, { coercionType: Office.CoercionType.Html }, done);
      };
      if (settings.disableOutlookSig && item.disableClientSignatureAsync) item.disableClientSignatureAsync(setSig);
      else setSig();
    };

    var fallback = mailbox.userProfile ? mailbox.userProfile.emailAddress : "";
    if (item.from && item.from.getAsync) {
      item.from.getAsync(function (r) {
        apply(r.status === Office.AsyncResultStatus.Succeeded && r.value ? r.value.emailAddress : fallback);
      });
    } else {
      apply(fallback);
    }
  } catch (e) {
    done();
  }
}

function onNewMessageComposeHandler(event) { nvInsertSignature(event); }
function onMessageFromChangedHandler(event) { nvInsertSignature(event); }

if (typeof Office !== "undefined" && Office.actions) {
  Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
  Office.actions.associate("onMessageFromChangedHandler", onMessageFromChangedHandler);
}
