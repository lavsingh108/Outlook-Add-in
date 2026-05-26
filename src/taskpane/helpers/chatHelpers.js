import { formatResponse } from "../utils/formatUtils.js";

export function showTypingIndicator() {
    const hist = document.getElementById("chat-history");
    if (hist.querySelector(".msg-typing")) return;
    const div = document.createElement("div");
    div.className = "msg-typing";
    div.id = "typing-indicator";
    div.innerHTML = `<span class="typing-dot"></span><span class="typing-dot"></span><span class="typing-dot"></span>`;
    hist.appendChild(div);
    hist.scrollTop = hist.scrollHeight;
}

export function hideTypingIndicator() {
    const el = document.getElementById("typing-indicator");
    if (el) el.remove();
}

export function appendMessage(role, text) {
    const hist = document.getElementById("chat-history");
    const div  = document.createElement("div");
    if (role === "user") {
        div.className = "msg-user";
        const p = document.createElement("p");
        p.textContent = text;
        div.appendChild(p);
    } else {
        div.className = "msg-ai";
        div.innerHTML = formatResponse(text);
    }
    hist.appendChild(div);
    hist.scrollTop = hist.scrollHeight;
}

export function hideSuggestions() {
    const box = document.getElementById("suggestions");
    box.classList.add("hidden");
    box.innerHTML = "";
}

// onSelect: called when the user clicks a suggestion chip (no args)
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

// onSelect: forwarded to renderSuggestions for suggestion chips
export function restoreConversationHistory(messages, onSelect) {
    messages.forEach(msg => {
        const text = msg.text || "";
        if (!text.trim()) return;
        appendMessage(msg.sender === "assistant" ? "ai" : "user", text);
    });

    const lastAI = [...messages].reverse().find(m => m.sender === "assistant");
    if (lastAI && Array.isArray(lastAI.tags) && lastAI.tags.length) {
        renderSuggestions(lastAI.tags, onSelect);
    }

    const hist = document.getElementById("chat-history");
    hist.scrollTop = hist.scrollHeight;
}
