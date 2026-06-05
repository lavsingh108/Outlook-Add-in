// ── Navigation ──────────────────────────────────────────────────────────────
// View-switching helpers shared between read and compose modes.

import { hideSuggestions } from "./ui/chatUI.js";
import {
    state, _chatFromCompose,
    setChatFromCompose,
} from "./state.js";
import { renderPreviousChats, loadReadAttachments } from "./readMode.js";

/**
 * Close the chat view and return to either compose or read.
 */
export function switchToReadView() {
    document.getElementById("view-chat").classList.add("hidden");

    // Return to compose if chat was opened from compose mode
    if (_chatFromCompose) {
        setChatFromCompose(false);
        document.getElementById("view-compose").classList.remove("hidden");
        document.getElementById("btn-back").classList.add("hidden");
        document.getElementById("chat-history").innerHTML = "";
        hideSuggestions();
        state.currentConversationId = null;
        state.currentDocumentId     = null;

        const startChatBtn = document.getElementById("btn-compose-start-chat");
        if (startChatBtn) startChatBtn.disabled = false;
        return;
    }

    // Return to read view
    document.getElementById("view-read").classList.remove("hidden");
    document.getElementById("btn-back").classList.add("hidden");
    document.getElementById("chat-history").innerHTML = "";
    hideSuggestions();
    state.currentConversationId = null;
    state.currentDocumentId     = null;

    // Re-enable share chat button (if present)
    const shareBtn = document.getElementById("btn-share-chat");
    if (shareBtn) shareBtn.disabled = false;

    renderPreviousChats();
    loadReadAttachments();

    const statusEl = document.getElementById("status-msg");
    if (statusEl) statusEl.innerText = "";
}
