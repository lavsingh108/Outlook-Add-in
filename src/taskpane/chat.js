// ── Chat Controller ──────────────────────────────────────────────────────────
// enterChat: open the chat view, restore or fetch welcome message.
// sendChatMessage: submit a user question to the ask API.

import { getAuthToken, clearToken } from "./services/authService.js";
import { fetchHistory, fetchWelcome, askQuestion } from "./services/apiService.js";
import {
    showTypingIndicator, hideTypingIndicator,
    appendMessage, renderSuggestions, hideSuggestions,
    restoreConversationHistory,
} from "./ui/chatUI.js";
import { state } from "./state.js";

/**
 * Open the chat view for a given conversation.
 * If the conversation already has assistant messages, history is restored
 * and the welcome API is skipped.
 */
export async function enterChat(conversationId, documentId, token) {
    state.currentConversationId = conversationId;
    state.currentDocumentId     = documentId;

    document.getElementById("view-read-init").classList.add("hidden");
    document.getElementById("view-read").classList.add("hidden");
    document.getElementById("view-compose").classList.add("hidden");
    document.getElementById("view-chat").classList.remove("hidden");
    document.getElementById("btn-back").classList.remove("hidden");
    document.getElementById("chat-history").innerHTML = "";
    hideSuggestions();

    if (!documentId) {
        appendMessage("ai", "Cannot start chat: document ID is missing. Please re-share the document.", conversationId, documentId);
        return;
    }

    showTypingIndicator();

    // Step 1: try to restore existing conversation history
    try {
        const histResp = await fetchHistory(token, conversationId);
        if (histResp.ok) {
            const conv  = await histResp.json();
            const msgs  = Array.isArray(conv.messages) ? conv.messages : [];
            const hasAI = msgs.some(m => m.sender === "assistant" || m.role === "assistant");
            if (hasAI) {
                hideTypingIndicator();
                restoreConversationHistory(msgs, conversationId, documentId, sendChatMessage);
                return;
            }
        }
    } catch (histErr) {
        console.warn("History fetch failed (non-fatal):", histErr.message);
    }

    // Step 2: no prior history — call welcome (first open only)
    try {
        const resp    = await fetchWelcome(token, conversationId, documentId);
        const rawText = await resp.text();
        hideTypingIndicator();

        if (!resp.ok) {
            console.error("Welcome API failed:", resp.status, rawText);
            appendMessage("ai", "Could not load welcome message (" + resp.status + "). You can still ask questions below.", conversationId, documentId);
            return;
        }

        let data;
        try { data = JSON.parse(rawText); }
        catch {
            appendMessage("ai", "Hello! How can I help you with this document?", conversationId, documentId);
            return;
        }

        const welcomeMsg =
            data.answer || data.response || data.message || data.text ||
            data.content || data.welcomeText || data.welcome_text ||
            (typeof data === "string" ? data : null);
        appendMessage("ai", welcomeMsg || "Hello! How can I help you with this document?", conversationId, documentId);

        const tags = Array.isArray(data.tags) ? data.tags : [];
        if (tags.length) renderSuggestions(tags, sendChatMessage);
    } catch (err) {
        hideTypingIndicator();
        console.error("enterChat welcome error:", err);
        appendMessage("ai", "Network error. You can still ask questions below.", conversationId, documentId);
    }
}

/**
 * Send the text from #user-input to the ask API and display the response.
 */
export async function sendChatMessage() {
    const input = document.getElementById("user-input");
    const text  = input.value.trim();
    if (!text) return;

    hideSuggestions();
    appendMessage("user", text, state.currentConversationId, state.currentDocumentId);
    input.value = "";
    document.getElementById("btn-send").disabled = true;
    showTypingIndicator();

    try {
        const token = await getAuthToken();
        const resp  = await askQuestion(token, state.currentConversationId, text);
        hideTypingIndicator();
        if (!resp.ok) throw new Error("Ask failed (" + resp.status + "): " + await resp.text());

        const data = await resp.json();
        appendMessage(
            "ai",
            data.answer || data.response || "No response received.",
            state.currentConversationId,
            state.currentDocumentId
        );

        const tags = Array.isArray(data.tags) ? data.tags : [];
        if (tags.length) renderSuggestions(tags, sendChatMessage);
    } catch (err) {
        hideTypingIndicator();
        appendMessage("ai", "Error: " + err.message, state.currentConversationId, state.currentDocumentId);
        clearToken();
    } finally {
        document.getElementById("btn-send").disabled = false;
    }
}
