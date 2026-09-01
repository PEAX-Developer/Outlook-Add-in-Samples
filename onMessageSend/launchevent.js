console.log("onMessageSend/launchevent.js loaded");

Office.initialize = function () {
}

const PIXEL_URL = new URL('/onMessageSend/assets/pixel.gif?source=on-message-send-repro', location.origin).href;

function getBodyHtml() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, result => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(result.error);
      }
    });
  });
}

function setBodyHtml(html) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.setAsync(
      html,
      { coercionType: Office.CoercionType.Html },
      result => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          reject(result.error);
        }
      }
    );
  });
}

function findQuotedReplyBoundary(document) {
  return document.querySelector(
    '[id$="divRplyFwdMsg"], a[name="_MailOriginal"], span[style*="_MailOriginalBody"]'
  );
}

function isBeforeBoundary(node, boundary) {
  return (
    !boundary ||
    Boolean(node.compareDocumentPosition(boundary) & Node.DOCUMENT_POSITION_FOLLOWING)
  );
}

// REPRO AID (not a fix): New Outlook's offline Action Queue evaluates the
// SendItem action's attachment "blocking keys" synchronously. If Send is
// allowed to proceed the instant setBodyHtml() resolves, the add-in
// sometimes races ahead of the platform registering the blocking key at
// all, so the platform bug only reproduces intermittently. Deliberately
// delaying event.completed() by a short, fixed amount reliably lands
// inside the platform's failure window instead, making
// ErrorLocalAttachmentBlockingKeys reproduce on (close to) every send -
// useful for capturing a clean HAR/console repro to attach to the platform
// bug escalation. This is the opposite of a mitigation: it does not wait
// on any real "attachment upload complete" signal (no such Office.js
// signal exists), it just widens the window in which the platform's race
// is lost. Do not ship this delay in a production add-in.
function deliberateReproDelay(ms = 1000) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

async function onMessageSendHandler(event) {
  try {
    const originalHtml = await getBodyHtml();
    const document = new DOMParser().parseFromString(originalHtml, 'text/html');
    const quotedReplyBoundary = findQuotedReplyBoundary(document);

    for (const link of document.querySelectorAll('a[href]')) {
      if (!isBeforeBoundary(link, quotedReplyBoundary)) {
        continue;
      }

      const href = link.getAttribute('href');
      if (!href || !/^https?:/i.test(href)) {
        continue;
      }

      const url = new URL(href);
      url.searchParams.set('repro-click-tracking', '1');
      link.setAttribute('href', url.href);
    }

    const pixel = document.createElement('img');
    pixel.src = PIXEL_URL;
    pixel.alt = '';
    pixel.width = 1;
    pixel.height = 1;
    pixel.style.width = '1px';
    pixel.style.height = '1px';

    if (quotedReplyBoundary) {
      quotedReplyBoundary.before(pixel);
    } else {
      document.body.append(pixel);
    }

    // The important operation is writing the complete body back, including
    // unchanged quoted HTML and its existing cid: image references. This is
    // a single full-body write (not two), matching the original add-in
    // behavior so there is no risk of duplicating content: Office.js has no
    // API to replace an arbitrary range, only to get/set the whole body or
    // insert at the current selection, so a full round trip is unavoidable
    // here given the link-rewrite requirement.
    await setBodyHtml(document.documentElement.outerHTML);

    // See deliberateReproDelay() above: this delay exists only to make the
    // platform race reproduce reliably for escalation purposes.
    await deliberateReproDelay();

    event.completed({ allowEvent: true });
  } catch (error) {
    console.error('Minimal onMessageSend repro failed', error);
    event.completed({
      allowEvent: false,
      errorMessage: 'The minimal onMessageSend repro handler failed.',
    });
  }
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);