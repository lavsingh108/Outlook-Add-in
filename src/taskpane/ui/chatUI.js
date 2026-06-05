// ── Chat UI ─────────────────────────────────────────────────────────────────
// DOM helpers scoped to the chat view: typing indicator, message bubbles,
// suggestion chips, and history restoration.

import { formatResponse } from "../utils/helpers.js";

// ── Typing indicator ───────────────────────────────────────────────────────

export function showTypingIndicator() {
    const hist = document.getElementById("chat-history");
    if (hist.querySelector(".msg-typing")) return;
    const div = document.createElement("div");
    div.className = "msg-typing";
    div.id        = "typing-indicator";
    div.innerHTML = `<span class="typing-dot"></span><span class="typing-dot"></span><span class="typing-dot"></span>`;
    hist.appendChild(div);
    hist.scrollTop = hist.scrollHeight;
}

export function hideTypingIndicator() {
    const el = document.getElementById("typing-indicator");
    if (el) el.remove();
}

// ── Messages ───────────────────────────────────────────────────────────────

export function appendMessage(role, text, conversationId, documentId) {
    const hist = document.getElementById("chat-history");
    const div  = document.createElement("div");

    if (role === "user") {
        div.className = "msg-user";
        const p = document.createElement("p");
        p.textContent = text;
        div.appendChild(p);
    } else {
        div.className = "msg-ai";
        div.innerHTML = formatResponse(text, conversationId, documentId);
    }

    hist.appendChild(div);
    hist.scrollTop = hist.scrollHeight;
}

// ── Suggestion chips ───────────────────────────────────────────────────────

export function hideSuggestions() {
    const box = document.getElementById("suggestions");
    box.classList.add("hidden");
    box.innerHTML = "";
}

export function renderSuggestions(tags, onSelect) {
    const box = document.getElementById("suggestions");
    box.innerHTML = "";

    tags.forEach(tag => {
        const q = typeof tag === "string" ? tag : (tag["next-question"] || tag.question || "");
        if (!q.trim()) return;

        const chip = document.createElement("button");
        chip.className = "chip";
        chip.textContent = q;
        chip.onclick = () => {
            hideSuggestions();
            document.getElementById("user-input").value = q;
            onSelect();
        };
        box.appendChild(chip);
    });

    box.classList.remove("hidden");
}

// ── History restoration ────────────────────────────────────────────────────

export function restoreConversationHistory(messages, conversationId, documentId, onSuggest) {
    messages.forEach(msg => {
        const text = msg.text || "";
        if (!text.trim()) return;
        appendMessage(msg.sender === "assistant" ? "ai" : "user", text, conversationId, documentId);
    });

    const lastAI = [...messages].reverse().find(m => m.sender === "assistant");
    if (lastAI && Array.isArray(lastAI.tags) && lastAI.tags.length) {
        renderSuggestions(lastAI.tags, onSuggest);
    }

    const hist = document.getElementById("chat-history");
    hist.scrollTop = hist.scrollHeight;
}

// ── Status messages ────────────────────────────────────────────────────────

export function showReadStatus(msg) {
    const el = document.getElementById("status-msg");
    if (el) el.innerText = msg;
}

export function showComposeStatus(msg) {
    document.getElementById("compose-status").innerText = msg;
}

export function showReadInitError(msg) {
    document.querySelector(".read-spinner-wrap").style.display = "none";
    document.getElementById("read-init-status").classList.add("hidden");
    document.getElementById("read-error-msg").textContent = msg;
    document.getElementById("read-init-error").classList.remove("hidden");
}
