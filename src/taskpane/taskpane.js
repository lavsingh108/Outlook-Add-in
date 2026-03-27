/* global Office, document, fetch, FormData, atob, Uint8Array, Blob, console */

/**
 * SmartBlue Outlook Add-in Logic
 * Migrated from Google Apps Script to Office.js
 */

// ── API Endpoints ─────────────────────────────────────────────────
const UPLOAD_URL     = "https://ws.demo.smartblue.ai/v1/document/upload";
const BUNDLE_ADD_URL = "https://ws.demo.smartblue.ai/v1/document/bundle/add";
const WELCOME_URL    = "https://ws.demo.smartblue.ai/v1/conversation/ask/welcome";
const ASK_URL        = "https://ws.demo.smartblue.ai/v1/conversation/ask/question";

// ── State ─────────────────────────────────────────────────────────
let currentConversationId = null;
let currentDocumentId     = null;
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

    // Event Listeners
    document.getElementById("chk-bundle").addEventListener("change", onBundleToggle);
    document.getElementById("btn-upload").addEventListener("click", handleUpload);
    document.getElementById("btn-send").addEventListener("click", sendChatMessage);
    document.getElementById("btn-retry").addEventListener("click", resetToAttachments);
    
    // Enter key to send
    document.getElementById("user-input").addEventListener("keydown", (e) => {
        if (e.key === "Enter" && !e.shiftKey) {
            e.preventDefault();
            sendChatMessage();
        }
    });
}

// ── Auth: Microsoft Office SSO (Native) ──────────────────────────
async function getAuthToken() {
    try {
        /**
         * Uses the built-in Office Identity API.
         * Requires <WebApplicationInfo> in manifest.xml
         */
        const token = await Office.auth.getAccessToken({
            allowSignInPrompt: true,
            allowConsentPrompt: true,
            forMSGraphAccess: false
        });
        console.log("Office SSO Token acquired");
        return token;
    } catch (err) {
        console.error("SSO Error:", err);
        // Error 13007: User needs to grant consent
        if (err.code === 13007) {
            throw new Error("Please grant permission to the add-in in the popup and try again.");
        }
        // Error 13000: Manifest configuration issue
        if (err.code === 13000) {
            throw new Error("Identity API not supported. Check manifest WebApplicationInfo.");
        }
        throw new Error("Authentication failed: " + err.message);
    }
}

// ── Load Attachments ──────────────────────────────────────────────
function loadAttachments() {
    const item = Office.context.mailbox.item;
    const attachments = item.attachments;
    const listDiv = document.getElementById("attachment-list");

    if (!attachments || attachments.length === 0) {
        listDiv.innerHTML = "<div class='empty-state'>No attachments found.</div>";
        document.getElementById("btn-upload").disabled = true;
        document.getElementById("bundle-toggle-row").style.display = "none";
        return;
    }

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

    renderPrimaryList(attachments);
}

function renderPrimaryList(attachments) {
    const primaryListDiv = document.getElementById("primary-list");
    const primaryLabel = document.getElementById("primary-label");
    
    primaryListDiv.innerHTML = "";

    if (!isBundleMode || attachments.length <= 1) {
        primaryLabel.style.display = "none";
        primaryListDiv.style.display = "none";
        return;
    }

    primaryLabel.style.display = "block";
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
    renderPrimaryList(Office.context.mailbox.item.attachments);
}

// ── Upload & Analyse Pipeline ─────────────────────────────────────
async function handleUpload() {
    const item = Office.context.mailbox.item;
    const attachments = item.attachments;

    showView("loading");
    setLoadingText("Authenticating...");

    try {
        const token = await getAuthToken();

        // 1. Upload Primary Document
        const pIdx = isBundleMode ? primaryIndex : 0;
        const primaryAtt = attachments[pIdx];

        setLoadingText(`Uploading ${primaryAtt.name}...`);
        const primaryBlob = await getAttachmentBlob(primaryAtt.id);
        const fd = new FormData();
        fd.append("document", primaryBlob, primaryAtt.name);

        const uploadResp = await fetch(UPLOAD_URL, {
            method: "POST",
            headers: { Authorization: `Bearer ${token}` },
            body: fd
        });

        if (!uploadResp.ok) throw new Error("Primary upload failed: " + uploadResp.status);

        const uploadData = await uploadResp.json();
        currentConversationId = uploadData.conversation_id || uploadData.conversationId;
        currentDocumentId     = uploadData.doc_id || uploadData.documentId;

        // 2. Upload Supporting Docs
        if (isBundleMode && attachments.length > 1) {
            for (let i = 0; i < attachments.length; i++) {
                if (i === pIdx) continue;
                setLoadingText(`Adding ${attachments[i].name} to bundle...`);
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

        // 3. Welcome Analysis
        setLoadingText("Analyzing documents...");
        const welcomeData = await callWelcomeAPI(token);

        showView("chat");
        if (welcomeData.message) appendMessage("ai", welcomeData.message);
        if (welcomeData.tags) renderSuggestions(welcomeData.tags);

    } catch (err) {
        showError(err.message);
    }
}

// ── Chat Functionality ────────────────────────────────────────────
async function sendChatMessage() {
    const input = document.getElementById("user-input");
    const text = input.value.trim();
    if (!text) return;

    input.value = "";
    hideSuggestions();
    appendMessage("user", text);
    
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
        const answer = data.answer || data.response || "No response.";
        
        appendMessage("ai", formatAIResponse(answer));
        if (data.tags) renderSuggestions(data.tags);

    } catch (err) {
        removeTypingIndicator(typingId);
        appendMessage("ai", "Error: " + err.message);
    }
}

// ── Helpers ───────────────────────────────────────────────────────

async function callWelcomeAPI(token) {
    const resp = await fetch(WELCOME_URL, {
        method: "POST",
        headers: { "Content-Type": "application/json", Authorization: `Bearer ${token}` },
        body: JSON.stringify({ conversationId: currentConversationId, documentId: currentDocumentId })
    });
    return resp.ok ? await resp.json() : { message: "Analysis complete. How can I help?" };
}

function getAttachmentBlob(attachmentId) {
    return new Promise((resolve, reject) => {
        Office.context.mailbox.item.getAttachmentContentAsync(attachmentId, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                const base64 = result.value.content;
                const binary = atob(base64);
                const bytes = new Uint8Array(binary.length);
                for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
                resolve(new Blob([bytes]));
            } else reject(new Error(result.error.message));
        });
    });
}

function formatAIResponse(text) {
    if (!text) return "";
    return text
        .replace(/\*\*(.*?)\*\*/g, "<b>$1</b>")
        .replace(/\*(.*?)\*/g, "<i>$1</i>")
        .replace(/\n/g, "<br>");
}

function renderSuggestions(tags) {
    const list = document.getElementById("suggestions-list");
    const bar = document.getElementById("suggestions-bar");
    list.innerHTML = "";
    tags.forEach(t => {
        const q = typeof t === 'string' ? t : (t['next-question'] || "");
        if (!q) return;
        const btn = document.createElement("button");
        btn.className = "suggestion-btn";
        btn.innerText = q;
        btn.onclick = () => {
            document.getElementById("user-input").value = q;
            sendChatMessage();
        };
        list.appendChild(btn);
    });
    bar.classList.remove("hidden");
}

function appendMessage(role, html) {
    const hist = document.getElementById("chat-history");
    const div = document.createElement("div");
    div.className = role === "user" ? "msg-row msg-user" : "msg-row msg-ai";
    div.innerHTML = `<div class="bubble">${html}</div>`;
    hist.appendChild(div);
    document.getElementById("chat-window").scrollTop = hist.scrollHeight;
}

function appendTypingIndicator() {
    const id = "typing-" + Date.now();
    const hist = document.getElementById("chat-history");
    const div = document.createElement("div");
    div.id = id;
    div.className = "msg-row msg-ai";
    div.innerHTML = `<div class="bubble">...</div>`;
    hist.appendChild(div);
    return id;
}

function removeTypingIndicator(id) {
    const el = document.getElementById(id);
    if (el) el.remove();
}

function hideSuggestions() { document.getElementById("suggestions-bar").classList.add("hidden"); }

function showView(name) {
    document.querySelectorAll(".view").forEach(v => v.classList.add("hidden"));
    document.getElementById("view-" + name).classList.remove("hidden");
}

function showError(msg) {
    document.getElementById("error-msg").innerText = msg;
    showView("error");
}

function resetToAttachments() { showView("attachments"); }

function setLoadingText(msg) { document.getElementById("loading-text").innerText = msg; }

function formatBytes(b) {
    if (b < 1024) return b + " B";
    return (b / 1024).toFixed(1) + " KB";
}

function getIconForName(n) {
    const ext = n.split('.').pop().toLowerCase();
    const icons = { pdf: "📕", docx: "📘", xlsx: "📗", pptx: "📙", png: "🖼️", jpg: "🖼️" };
    return icons[ext] || "📄";
}