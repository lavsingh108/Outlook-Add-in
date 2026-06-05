// ── Entry Point ─────────────────────────────────────────────────────────────
// Bootstraps the add-in once Office is ready.

import { initRead }    from "./readMode.js";
import { initCompose } from "./composeMode.js";

Office.onReady(info => {
    if (info.host === Office.HostType.Outlook) init();
});

function init() {
    const item      = Office.context.mailbox.item;
    const isCompose = typeof item.subject?.setAsync === "function"
                   || typeof item.body?.setAsync    === "function";
    if (isCompose) initCompose();
    else           initRead();
}
