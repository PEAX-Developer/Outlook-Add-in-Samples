console.log("onMessageSend/launchevent.js loaded");

Office.initialize = function () {
}

// Any publicly reachable image works here - this is just a well-known,
// always-available URL for repro purposes. Nothing about this URL is an
// attachment; it's a plain remote image reference.
const BANNER_IMG_URL = "https://img-prod-cms-rt-microsoft-com.akamaized.net/cms/api/am/imageFileData/RE1Mu3b";

// Minimal onMessageSend repro: prepend a one-line HTML banner (a single
// external <img>) to the existing body and write the whole body back.
// No attachments, links, or existing message content are read or touched
// beyond this. On reply, this alone is enough to trigger a send failure in
// New Outlook.
async function onMessageSendHandler(event) {
  Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, getResult => {
    if (getResult.status !== Office.AsyncResultStatus.Succeeded) {
      event.completed({
        allowEvent: false,
        errorMessage: "Could not read the message body.",
      });
      return;
    }

    const banner = `<div><img src="${BANNER_IMG_URL}" alt="Banner" /></div>`;
    const newBody = banner + getResult.value;

    Office.context.mailbox.item.body.setAsync(
      newBody,
      { coercionType: Office.CoercionType.Html },
      setResult => {
        if (setResult.status !== Office.AsyncResultStatus.Succeeded) {
          event.completed({
            allowEvent: false,
            errorMessage: "Could not update the message body.",
          });
          return;
        }

        event.completed({ allowEvent: true });
      }
    );
  });
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
