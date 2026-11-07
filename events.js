/* Automatisch invoegen: draait bij een nieuwe mail en wanneer je het Van-adres wijzigt. */
var NV_EVENTS_VERSION = "1.0.3";

// followFrom: bij een gewijzigd Van-adres altijd de versie van dat adres kiezen,
// ook als in het paneel een vaste standaardversie is opgeslagen.
function nvInsertSignature(event, followFrom) {
  var finished = false;
  var done = function () { if (finished) return; finished = true; try { event.completed(); } catch (e) {} };
  try {
    var mailbox = Office.context.mailbox;
    var item = mailbox.item;
    var settings = NVSig.load(Office.context.roamingSettings);
    if (!settings.autoInsert) { done(); return; }
    if (followFrom) settings.defaultVariant = "auto";

    var tries = 0;
    var setSig = function (html) {
      item.body.setSignatureAsync(html, { coercionType: Office.CoercionType.Html }, function (r) {
        // Direct na het wisselen van account is het bericht soms nog niet klaar. Dan een keer opnieuw.
        if (r.status === Office.AsyncResultStatus.Failed && tries < 2) { tries++; setTimeout(function () { setSig(html); }, 400 * tries); return; }
        done();
      });
    };

    var apply = function (fromAddress) {
      var html = NVSig.build(NVSig.pickVariant(settings, fromAddress), settings);
      if (settings.disableOutlookSig && item.disableClientSignatureAsync) item.disableClientSignatureAsync(function () { setSig(html); });
      else setSig(html);
    };

    var fallback = mailbox.userProfile ? mailbox.userProfile.emailAddress : "";
    var getFrom = function (attempt) {
      if (!(item.from && item.from.getAsync)) { apply(fallback); return; }
      item.from.getAsync(function (r) {
        var addr = r.status === Office.AsyncResultStatus.Succeeded && r.value ? r.value.emailAddress : "";
        if (!addr && attempt < 2) { setTimeout(function () { getFrom(attempt + 1); }, 300); return; }
        apply(addr || fallback);
      });
    };
    getFrom(0);
  } catch (e) {
    done();
  }
}

function onNewMessageComposeHandler(event) { nvInsertSignature(event, false); }
function onMessageFromChangedHandler(event) { nvInsertSignature(event, true); }

if (typeof Office !== "undefined" && Office.actions) {
  Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
  Office.actions.associate("onMessageFromChangedHandler", onMessageFromChangedHandler);
}
