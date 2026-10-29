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
