// ── Entry Point ─────────────────────────────────────────────────────────────
// Bootstraps the add-in once Office is ready.

import { initRead }    from "./readMode.js";
import { initCompose } from "./composeMode.js";
import { initSent }    from "./sentMode.js";

Office.onReady(info => {
    if (info.host === Office.HostType.Outlook) init();
});

function init() {
    const item      = Office.context.mailbox.item;
    const isCompose = typeof item.subject?.setAsync === "function"
                   || typeof item.body?.setAsync    === "function";

    if (isCompose) {
        initCompose();
        return;
    }

    // Detect sent items folder: Office.MailboxEnums.MessageType is not always
    // available, so we fall back to checking the sender address against the
    // current user's mailbox address.
    const isSent = _isSentItem(item);
    if (isSent) initSent();
    else        initRead();
}

/**
 * Heuristic to decide whether the currently open message lives in a Sent folder.
 *
 * Office.js exposes `item.displayType` or `item.messageClass` on some builds.
 * The most reliable cross-platform check is comparing the From address to the
 * current-user mailbox address — a sent message's From is the logged-in user.
 */
function _isSentItem(item) {
    try {
        // Preferred: Office 1.7+ exposes a dedicated property
        if (typeof item.getComposeTypeAsync === "undefined") {
            // `from` is only available on read items; if sender equals current user it's sent
            const userEmail = (Office.context.mailbox.userProfile?.emailAddress || "").toLowerCase();
            if (!userEmail) return false;

            // item.from is an EmailAddressDetails object in read mode
            const fromEmail = (item.from?.emailAddress || "").toLowerCase();
            return fromEmail !== "" && fromEmail === userEmail;
        }
    } catch (_) {}
    return false;
}
