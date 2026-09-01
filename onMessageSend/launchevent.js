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

// Insert HTML at the current cursor/selection (defaults to the start of the
// compose body) instead of getAsync()+setAsync() of the whole document. This
// never touches or re-serializes the existing quoted thread/cid: images, so
// New Outlook has nothing extra to reconcile there when Send is pressed.
function insertAtStart(html) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.setSelectedDataAsync(
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

// Give New Outlook a bounded window to resolve any newly-inserted inline
// image's blob id to a server content id before letting Send proceed. This
// is a defensive backstop for the ErrorLocalAttachmentBlockingKeys race, not
// a guarantee - Office.js has no "attachment upload complete" event.
function waitForAttachmentsSettled({ retries = 6, delayMs = 250 } = {}) {
  return new Promise(resolve => {
    let attempt = 0;
    let lastCount = -1;

    const check = () => {
      Office.context.mailbox.item.getAttachmentsAsync(result => {
        const count =
          result.status === Office.AsyncResultStatus.Succeeded
            ? result.value.length
            : lastCount;

        attempt += 1;
        if (count === lastCount || attempt >= retries) {
          resolve();
          return;
        }
        lastCount = count;
        setTimeout(check, delayMs);
      });
    };

    check();
  });
}

async function onMessageSendHandler(event) {
  try {
    const originalHtml = await getBodyHtml();
    const document = new DOMParser().parseFromString(originalHtml, 'text/html');
    const quotedReplyBoundary = findQuotedReplyBoundary(document);

    // Only rewrite links that live in the new reply, above the quoted
    // boundary - never re-touch the quoted HTML itself.
    let linksRewritten = false;
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
      linksRewritten = true;
    }

    if (linksRewritten) {
      // Rewritten links only exist in the new-reply portion, so replace just
      // that fragment instead of round-tripping the whole document.
      const newReplyFragment = [];
      for (const node of Array.from(document.body.childNodes)) {
        if (!isBeforeBoundary(node, quotedReplyBoundary)) {
          break;
        }
        newReplyFragment.push(node.outerHTML ?? node.textContent ?? '');
      }
      await insertAtStart(newReplyFragment.join(''));
    }

    // Insert the pixel via the targeted API too, so it is the only new
    // inline resource New Outlook needs to resolve - the quoted thread's
    // existing cid: images are never re-serialized or re-written.
    await insertAtStart(
      `<img src="${PIXEL_URL}" alt="" width="1" height="1" style="width:1px;height:1px;" />`
    );

    // Give New Outlook a bounded window to resolve the new inline image's
    // blob id to a server content id before we let Send proceed.
    await waitForAttachmentsSettled();

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