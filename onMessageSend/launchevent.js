console.log("onMessageSend/launchevent.js loaded");

Office.initialize = function () {
}

// The tracking pixel must stay a remote URL (not a file attachment) - the
// whole point is that the recipient's mail client fetches it from our
// server, which is how the "open" event is recorded. A plain remote <img>
// is not an attachment link, so adding it to the body is fine; it's only
// *existing* attachment/cid: links elsewhere in the body that must not be
// touched by a plain body edit (see QUOTED_REPLY_BOUNDARY_PATTERN below).
const PIXEL_URL = new URL('/onMessageSend/assets/pixel.gif?source=on-message-send', location.origin).href;

// Matches the start of Outlook's quoted reply/forward markup. Everything at
// or after this index is left completely untouched (see rewriteHtml below)
// because it commonly contains existing inline image attachments referenced
// by cid:, and New Outlook does not support editing a body that contains
// attachment links except via the attachment APIs - re-serializing that
// markup (even without changing it) through a DOM parser is what causes
// ErrorLocalAttachmentBlockingKeys. Working on the raw HTML string instead
// guarantees byte-for-byte preservation of anything we don't explicitly
// touch.
const QUOTED_REPLY_BOUNDARY_PATTERN = /<[^>]*\bid=(["'])[^"']*divRplyFwdMsg\1[^>]*>|<a\b[^>]*\bname=(["'])_MailOriginal\2[^>]*>|<span\b[^>]*\bstyle=(["'])[^"']*_MailOriginalBody[^"']*\3[^>]*>/i;

const HREF_PATTERN = /(<a\b[^>]*\bhref=)(["'])(https?:\/\/[^"']*)\2/gi;

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

// Note: the tracking pixel intentionally stays a remote <img>, not a file
// attachment - if an add-in ever needs to add a genuinely new *inline*
// image (e.g. a signature logo), that must go through the attachment API:
// Office.context.mailbox.item.addFileAttachmentFromBase64Async(base64,
// fileName, { isInline: true }, callback), referencing the result in the
// body as <img src="cid:fileName">.

function rewriteHrefsForClickTracking(html) {
  return html.replace(HREF_PATTERN, (match, prefix, quote, href) => {
    try {
      const url = new URL(href);
      url.searchParams.set('click-tracking', '1');
      return `${prefix}${quote}${url.href}${quote}`;
    } catch {
      // Not a parseable absolute URL; leave the link exactly as-is.
      return match;
    }
  });
}

async function onMessageSendHandler(event) {
  try {
    const originalHtml = await getBodyHtml();

    const boundaryMatch = originalHtml.match(QUOTED_REPLY_BOUNDARY_PATTERN);
    const boundaryIndex = boundaryMatch ? boundaryMatch.index : originalHtml.length;

    // headHtml is the only part of the body we ever rewrite. tailHtml
    // (the quoted/forwarded history and anything in it, including existing
    // cid: attachment references) is carried through unchanged, character
    // for character.
    const headHtml = originalHtml.slice(0, boundaryIndex);
    const tailHtml = originalHtml.slice(boundaryIndex);

    const rewrittenHead = rewriteHrefsForClickTracking(headHtml);
    const pixelTag = `<img src="${PIXEL_URL}" alt="" width="1" height="1" style="width:1px;height:1px;" />`;

    // Single full-body write (Office.js has no API to replace only part of
    // the body), but everything after headHtml is untouched original bytes -
    // no existing attachment/cid: reference anywhere in the body is
    // re-serialized or otherwise touched.
    await setBodyHtml(rewrittenHead + pixelTag + tailHtml);

    event.completed({ allowEvent: true });
  } catch (error) {
    console.error('onMessageSendHandler failed', error);
    event.completed({
      allowEvent: false,
      errorMessage: 'The onMessageSend handler failed.',
    });
  }
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);