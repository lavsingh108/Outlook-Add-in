// ── Navigation ──────────────────────────────────────────────────────────────
// View-switching helpers shared between read, sent, and compose modes.

import { hideSuggestions } from "./ui/chatUI.js";
import {
    state, _chatFromCompose, _chatFromSent,
    setChatFromCompose, setChatFromSent,
} from "./state.js";
import { renderPreviousChats, loadReadAttachments } from "./readMode.js";

/**
 * Close the chat view and return to the view that opened it
 * (compose → compose, sent → sent, otherwise → read).
 */
export function switchToReadView() {
    document.getElementById("view-chat").classList.add("hidden");
    document.getElementById("chat-history").innerHTML = "";
    hideSuggestions();
    state.currentConversationId = null;
    state.currentDocumentId     = null;

    // Return to compose if chat was opened from compose mode
    if (_chatFromCompose) {
        setChatFromCompose(false);
        document.getElementById("view-compose").classList.remove("hidden");
        document.getElementById("btn-back").classList.add("hidden");

        const startChatBtn = document.getElementById("btn-compose-start-chat");
        if (startChatBtn) startChatBtn.disabled = false;
        return;
    }

    // Return to sent view if chat was opened from sent mode
    if (_chatFromSent) {
        setChatFromSent(false);
        document.getElementById("view-sent").classList.remove("hidden");
        document.getElementById("btn-back").classList.add("hidden");
        const sentStatus = document.getElementById("sent-status-msg");
        if (sentStatus) sentStatus.innerText = "";
        return;
    }

    // Default: return to read view
    document.getElementById("view-read").classList.remove("hidden");
    document.getElementById("btn-back").classList.add("hidden");

    // Re-enable share chat button (if present)
    const shareBtn = document.getElementById("btn-share-chat");
    if (shareBtn) shareBtn.disabled = false;

    renderPreviousChats();
    loadReadAttachments();

    const statusEl = document.getElementById("status-msg");
    if (statusEl) statusEl.innerText = "";
}
