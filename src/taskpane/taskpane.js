/* global Office, document, fetch, FormData, atob, Uint8Array, Blob, console */

// ── API Endpoints ─────────────────────────────────────────────────
const UPLOAD_URL     = "https://ws.demo.smartblue.ai/v1/document/upload";
const BUNDLE_ADD_URL = "https://ws.demo.smartblue.ai/v1/document/bundle/add";
const WELCOME_URL    = "https://ws.demo.smartblue.ai/v1/conversation/ask/welcome";
const ASK_URL        = "https://ws.demo.smartblue.ai/v1/conversation/ask/question";

// ── State ─────────────────────────────────────────────────────────
let currentConversationId = null;
let currentDocumentId     = null;
let chatHistory           = [];  // { role: "user"|"ai", text: string }[]
let isBundleMode          = true;
let primaryIndex          = 0;

// ── Init ──────────────────────────────────────────────────────────
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        init();
    }
});

function init() {
    loadAttachments();

    document.getElementById("chk-bundle").addEventListener("change", onBundleToggle);
    document.getElementById("btn-upload").addEventListener("click", handleUpload);
    document.getElementById("btn-send").addEventListener("click", sendChatMessage);
    document.getElementById("user-input").addEventListener("keydown", (e) => {
        if (e.key === "Enter" && !e.shiftKey) {
            e.preventDefault();
            sendChatMessage();
        }
    });
}

// ── Auth: Microsoft Entra (Office SSO) ───────────────────────────
async function getAuthToken() {
    try {
        // Gets the Microsoft Entra access token via Office SSO
        const token = await Office.auth.getAccessToken({
            allowSignInPrompt: true,
            allowConsentPrompt: true,
            forMSGraphAccess: false
        });
        return token;
    } catch (err) {
        console.error("Entra auth error:", err);
        throw new Error("Authentication failed: " + (err.message || err.code));
    }
}

// ── Load Attachments ──────────────────────────────────────────────
function loadAttachments() {
    const item = Office.context.mailbox.item;
    const attachments = item.attachments;
    const listDiv = document.getElementById("attachment-list");
    const primaryListDiv = document.getElementById("primary-list");

    if (!attachments || attachments.length === 0) {
        listDiv.innerHTML = "<div class='empty-state'>No attachments found in this email.</div>";
        document.getElementById("btn-upload").disabled = true;
        document.getElementById("bundle-toggle-row").style.display = "none";
        document.getElementById("primary-label").style.display = "none";
        return;
    }

    // Attachment info list
    listDiv.innerHTML = "";
    attachments.forEach((att) => {
        const div = document.createElement("div");
        div.className = "att-card";
        div.innerHTML = `
            <span class="att-icon">${getIconForName(att.name)}</span>
            <div class="att-info">
                <div class="att-name">${att.name}</div>
                <div class="att-size">${formatBytes(att.size)}</div>
            </div>`;
        listDiv.appendChild(div);
    });

    // Primary radio list
    renderPrimaryList(attachments);
}

function renderPrimaryList(attachments) {
    const primaryListDiv = document.getElementById("primary-list");
    primaryListDiv.innerHTML = "";

    if (!isBundleMode || attachments.length <= 1) {
        document.getElementById("primary-label").style.display = "none";
        primaryListDiv.style.display = "none";
        return;
    }

    document.getElementById("primary-label").style.display = "block";
    primaryListDiv.style.display = "block";

    attachments.forEach((att, i) => {
        const div = document.createElement("div");
        div.className = "radio-item";
        div.innerHTML = `
            <label style="display:flex;align-items:center;gap:8px;cursor:pointer;">
                <input type="radio" name="primary" value="${i}" ${i === primaryIndex ? "checked" : ""}/>
                <span class="att-name-sm">${att.name}</span>
            </label>`;
        div.querySelector("input").addEventListener("change", () => { primaryIndex = i; });
        primaryListDiv.appendChild(div);
    });
}

function onBundleToggle() {
    isBundleMode = document.getElementById("chk-bundle").checked;
    const item = Office.context.mailbox.item;
    renderPrimaryList(item.attachments);
}

// ── Upload Handler ────────────────────────────────────────────────
async function handleUpload() {
    const item = Office.context.mailbox.item;
    const attachments = item.attachments;

    if (!attachments || attachments.length === 0) {
        showError("No attachments found.");
        return;
    }

    showView("loading");
    setLoadingText("Authenticating…");

    let token;
    try {
        token = await getAuthToken();
    } catch (err) {
        showError(err.message);
        return;
    }

    try {
        // Step 1: Upload primary document
        const pIdx = isBundleMode ? primaryIndex : 0;
        const primaryAtt = attachments[pIdx];

        setLoadingText(`Uploading ${primaryAtt.name}…`);
        const primaryBlob = await getAttachmentBlob(primaryAtt.id);
        const fd = new FormData();
        fd.append("document", primaryBlob, primaryAtt.name);

        const uploadResp = await fetch(UPLOAD_URL, {
            method: "POST",
            headers: { Authorization: `Bearer ${token}` },
            body: fd
        });

        if (!uploadResp.ok) {
            throw new Error(`Upload failed: HTTP ${uploadResp.status}`);
        }

        const uploadData = await uploadResp.json();
        currentConversationId = uploadData.conversation_id || uploadData.conversationId;
        currentDocumentId     = uploadData.doc_id || uploadData.documentId;

        if (!currentConversationId) throw new Error("No conversation_id returned from upload.");

        // Step 2: Upload supporting documents (bundle mode)
        if (isBundleMode && attachments.length > 1) {
            for (let i = 0; i < attachments.length; i++) {
                if (i === pIdx) continue;
                setLoadingText(`Uploading supporting: ${attachments[i].name}…`);
                const blob = await getAttachmentBlob(attachments[i].id);
                const sf = new FormData();
                sf.append("document", blob, attachments[i].name);
                await fetch(`${BUNDLE_ADD_URL}?conversation_id=${currentConversationId}`, {
                    method: "POST",
                    headers: { Authorization: `Bearer ${token}` },
                    body: sf
                });
            }
        }

        // Step 3: Welcome API
        setLoadingText("Preparing analysis…");
        const welcomeResult = await callWelcomeAPI(token);

        // Step 4: Switch to chat
        chatHistory = [];
        showView("chat");

        // Show welcome message
        if (welcomeResult.message) {
            appendMessage("ai", welcomeResult.message);
        }

        // Show suggested questions
        if (welcomeResult.tags && welcomeResult.tags.length > 0) {
            renderSuggestions(welcomeResult.tags);
        }

    } catch (err) {
        console.error("Upload error:", err);
        showError(err.message);
    }
}

// ── Welcome API ───────────────────────────────────────────────────
async function callWelcomeAPI(token) {
    try {
        const resp = await fetch(WELCOME_URL, {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
                Authorization: `Bearer ${token}`
            },
            body: JSON.stringify({
                conversationId: currentConversationId,
                documentId: currentDocumentId
            })
        });

        if (!resp.ok) return { message: "How can I help you today?", tags: [] };

        const data = await resp.json();
        return {
            message: data.message || data.answer || data.response || "How can I help you today?",
            tags: data.tags || []
        };
    } catch (err) {
        console.error("Welcome API error:", err);
        return { message: "How can I help you today?", tags: [] };
    }
}

// ── Chat ──────────────────────────────────────────────────────────
async function sendChatMessage() {
    const input = document.getElementById("user-input");
    const text = input.value.trim();
    if (!text) return;

    input.value = "";
    autoResizeTextarea(input);
    hideSuggestions();
    appendMessage("user", text);
    document.getElementById("btn-send").disabled = true;

    // Add typing indicator
    const typingId = appendTypingIndicator();

    try {
        const token = await getAuthToken();
        const resp = await fetch(ASK_URL, {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
                Authorization: `Bearer ${token}`
            },
            body: JSON.stringify({
                conversationId: currentConversationId,
                text: text,
                isMobile: false
            })
        });

        removeTypingIndicator(typingId);

        const data = await resp.json();
        const answer = data.answer || data.response || data.message || "No response received.";

        chatHistory.push({ role: "user", text });
        chatHistory.push({ role: "ai", text: answer });

        appendMessage("ai", formatAIResponse(answer));

        if (data.tags && data.tags.length > 0) {
            renderSuggestions(data.tags);
        }

    } catch (err) {
        removeTypingIndicator(typingId);
        appendMessage("ai", "Error: " + err.message);
    } finally {
        document.getElementById("btn-send").disabled = false;
    }
}

async function onSuggestedQuestion(question) {
    hideSuggestions();
    document.getElementById("user-input").value = question;
    await sendChatMessage();
}

// ── Suggested Questions ───────────────────────────────────────────
function renderSuggestions(tags) {
    const bar = document.getElementById("suggestions-bar");
    const list = document.getElementById("suggestions-list");
    list.innerHTML = "";

    tags.forEach((tag) => {
        const text = typeof tag === "string" ? tag : (tag["next-question"] || "");
        if (!text.trim()) return;
        const btn = document.createElement("button");
        btn.className = "suggestion-btn";
        btn.textContent = text;
        btn.addEventListener("click", () => onSuggestedQuestion(text));
        list.appendChild(btn);
    });

    bar.classList.remove("hidden");
}

function hideSuggestions() {
    document.getElementById("suggestions-bar").classList.add("hidden");
    document.getElementById("suggestions-list").innerHTML = "";
}

// ── Message Rendering ─────────────────────────────────────────────
function appendMessage(role, html) {
    const hist = document.getElementById("chat-history");
    const wrap = document.createElement("div");
    wrap.className = role === "user" ? "msg-row msg-user-row" : "msg-row msg-ai-row";

    if (role === "ai") {
        wrap.innerHTML = `
            <div class="ai-avatar">AI</div>
            <div class="bubble msg-ai-bubble">${html}</div>`;
    } else {
        wrap.innerHTML = `<div class="bubble msg-user-bubble">${html}</div>`;
    }

    hist.appendChild(wrap);
    scrollChatToBottom();
    return wrap;
}

function appendTypingIndicator() {
    const hist = document.getElementById("chat-history");
    const id = "typing-" + Date.now();
    const div = document.createElement("div");
    div.id = id;
    div.className = "msg-row msg-ai-row";
    div.innerHTML = `<div class="ai-avatar">AI</div><div class="bubble msg-ai-bubble typing-indicator"><span></span><span></span><span></span></div>`;
    hist.appendChild(div);
    scrollChatToBottom();
    return id;
}

function removeTypingIndicator(id) {
    const el = document.getElementById(id);
    if (el) el.remove();
}

function scrollChatToBottom() {
    const win = document.getElementById("chat-window");
    win.scrollTop = win.scrollHeight;
}

// ── AI Response Formatter (mirrors Gmail add-on formatAIResponse) ─
function formatAIResponse(text) {
    if (!text) return "";
    return text
        .replace(/\*\*(.*?)\*\*/g, "<b>$1</b>")
        .replace(/__(.*?)__/g, "<b>$1</b>")
        .replace(/\*(.*?)\*/g, "<i>$1</i>")
        .replace(/_(.*?)_/g, "<i>$1</i>")
        .replace(/^### (.*?)$/gm, "<b>$1</b>")
        .replace(/^## (.*?)$/gm, "<b>$1</b>")
        .replace(/^# (.*?)$/gm, "<b>$1</b>")
        .replace(/^[-*] (.*?)$/gm, "• $1")
        .replace(/^\d+\. (.*?)$/gm, "• $1")
        .replace(/\n/g, "<br>");
}

// ── Get Attachment as Blob ────────────────────────────────────────
function getAttachmentBlob(attachmentId) {
    return new Promise((resolve, reject) => {
        Office.context.mailbox.item.getAttachmentContentAsync(attachmentId, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                const base64 = result.value.content;
                const byteChars = atob(base64);
                const byteArr = new Uint8Array(byteChars.length);
                for (let i = 0; i < byteChars.length; i++) {
                    byteArr[i] = byteChars.charCodeAt(i);
                }
                resolve(new Blob([byteArr]));
            } else {
                reject(new Error(result.error.message));
            }
        });
    });
}

// ── View Management ───────────────────────────────────────────────
function showView(name) {
    ["view-attachments", "view-loading", "view-chat", "view-error"].forEach((id) => {
        document.getElementById(id).classList.add("hidden");
    });
    document.getElementById("view-" + name).classList.remove("hidden");
}

function showError(msg) {
    document.getElementById("error-msg").textContent = msg || "Unknown error.";
    showView("error");
}

function resetToAttachments() {
    currentConversationId = null;
    currentDocumentId = null;
    chatHistory = [];
    document.getElementById("chat-history").innerHTML = "";
    hideSuggestions();
    showView("attachments");
}

function setLoadingText(msg) {
    document.getElementById("loading-text").textContent = msg;
}

// ── Textarea Auto-Resize ──────────────────────────────────────────
function autoResizeTextarea(el) {
    el.style.height = "auto";
    el.style.height = Math.min(el.scrollHeight, 100) + "px";
}

document.addEventListener("DOMContentLoaded", () => {
    const ta = document.getElementById("user-input");
    if (ta) ta.addEventListener("input", () => autoResizeTextarea(ta));
});

// ── Helpers ───────────────────────────────────────────────────────
function formatBytes(bytes) {
    if (!bytes) return "";
    if (bytes < 1024) return bytes + " B";
    if (bytes < 1048576) return (bytes / 1024).toFixed(1) + " KB";
    return (bytes / 1048576).toFixed(1) + " MB";
}

function getIconForName(name) {
    if (!name) return "📄";
    const ext = name.split(".").pop().toLowerCase();
    const map = { pdf: "📕", doc: "📘", docx: "📘", xls: "📗", xlsx: "📗",
                  ppt: "📙", pptx: "📙", jpg: "🖼", jpeg: "🖼", png: "🖼",
                  zip: "📦", rar: "📦", mp4: "🎬", mp3: "🎵" };
    return map[ext] || "📄";
}
