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