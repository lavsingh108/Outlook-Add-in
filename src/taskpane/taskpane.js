/* global Office, document, fetch, FormData, atob, Uint8Array, Blob, console, msal */

// ── API Endpoints ─────────────────────────────────────────────────
const UPLOAD_URL     = "https://ws.demo.smartblue.ai/v1/document/upload";
const BUNDLE_ADD_URL = "https://ws.demo.smartblue.ai/v1/document/bundle/add";
const WELCOME_URL    = "https://ws.demo.smartblue.ai/v1/conversation/ask/welcome";
const ASK_URL        = "https://ws.demo.smartblue.ai/v1/conversation/ask/question";

// ── MSAL Config ───────────────────────────────────────────────────
// ⚠️ Replace these with your Azure App Registration values
const AZURE_CLIENT_ID = "YOUR_CLIENT_ID";  // Azure Portal → App registrations → Application (client) ID
const AZURE_TENANT_ID = "YOUR_TENANT_ID";  // Azure Portal → App registrations → Directory (tenant) ID

const msalConfig = {
    auth: {
        clientId:    AZURE_CLIENT_ID,
        authority:   `https://login.microsoftonline.com/${AZURE_TENANT_ID}`,
        redirectUri: window.location.origin + window.location.pathname
    },
    cache: {
        cacheLocation:        "sessionStorage",
        storeAuthStateInCookie: false
    }
};

const MSAL_SCOPES = ["openid", "profile", "email", "User.Read"];

let _msalInstance = null;
function getMsal() {
    if (!_msalInstance) _msalInstance = new msal.PublicClientApplication(msalConfig);
    return _msalInstance;
}

// ── State ─────────────────────────────────────────────────────────
let currentConversationId = null;
let currentDocumentId     = null;
let isBundleMode          = true;
let primaryIndex          = 0;

// ── Office Ready ──────────────────────────────────────────────────
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
    document.getElementById("btn-retry").addEventListener("click", resetToAttachments);
    document.getElementById("user-input").addEventListener("keydown", (e) => {
        if (e.key === "Enter" && !e.shiftKey) {
            e.preventDefault();
            sendChatMessage();
        }
    });
    // Auto-resize textarea
    document.getElementById("user-input").addEventListener("input", (e) => {
        e.target.style.height = "auto";
        e.target.style.height = Math.min(e.target.scrollHeight, 100) + "px";
    });
}

// ── Auth: MSAL.js (Microsoft Entra) ──────────────────────────────
async function getAuthToken() {
    const msalInstance = getMsal();

    // 1. Try silent token (uses cached session — no popup shown)
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
        try {
            const silent = await msalInstance.acquireTokenSilent({
                scopes:  MSAL_SCOPES,
                account: accounts[0]
            });
            console.log("MSAL silent token OK:", accounts[0].username);
            return silent.accessToken;
        } catch (silentErr) {
            console.warn("Silent token failed, trying popup:", silentErr.message);
        }
    }

    // 2. Popup login (first time or session expired)
    try {
        const popup = await msalInstance.loginPopup({
            scopes: MSAL_SCOPES,
            prompt: "select_account"
        });
        console.log("MSAL popup login OK:", popup.account.username);
        return popup.accessToken;
    } catch (popupErr) {
        console.error("MSAL popup error:", popupErr);
        throw new Error("Sign-in failed: " + (popupErr.message || popupErr.errorCode));
    }
}

// ── Load Attachments ──────────────────────────────────────────────
function loadAttachments() {
    const item        = Office.context.mailbox.item;
    const attachments = item.attachments;
    const listDiv     = document.getElementById("attachment-list");

    if (!attachments || attachments.length === 0) {
        listDiv.innerHTML = "<div class='empty-state'>No attachments found in this email.</div>";
        document.getElementById("btn-upload").disabled = true;
        document.getElementById("bundle-toggle-row").style.display = "none";
        document.getElementById("primary-label").style.display = "none";
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
    const primaryLabel   = document.getElementById("primary-label");
    primaryListDiv.innerHTML = "";

    if (!isBundleMode || attachments.length <= 1) {
        primaryLabel.style.display   = "none";
        primaryListDiv.style.display = "none";
        return;
    }

    primaryLabel.style.display   = "block";
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

// ── Upload Pipeline ───────────────────────────────────────────────
async function handleUpload() {
    const item        = Office.context.mailbox.item;
    const attachments = item.attachments;

    showView("loading");
    setLoadingText("Signing in…");

    try {
        const token = await getAuthToken();

        // 1. Upload primary document
        const pIdx       = isBundleMode ? primaryIndex : 0;
        const primaryAtt = attachments[pIdx];

        setLoadingText(`Uploading ${primaryAtt.name}…`);
        const primaryBlob = await getAttachmentBlob(primaryAtt.id);
        const fd = new FormData();
        fd.append("document", primaryBlob, primaryAtt.name);

        const uploadResp = await fetch(UPLOAD_URL, {
            method:  "POST",
            headers: { Authorization: `Bearer ${token}` },
            body:    fd
        });

        if (!uploadResp.ok) throw new Error("Upload failed: HTTP " + uploadResp.status);

        const uploadData      = await uploadResp.json();
        currentConversationId = uploadData.conversation_id || uploadData.conversationId;
        currentDocumentId     = uploadData.doc_id || uploadData.documentId;

        if (!currentConversationId) throw new Error("No conversation ID returned from server.");

        // 2. Upload supporting docs (bundle mode only)
        if (isBundleMode && attachments.length > 1) {
            for (let i = 0; i < attachments.length; i++) {
                if (i === pIdx) continue;
                setLoadingText(`Adding ${attachments[i].name} to bundle…`);
                const blob = await getAttachmentBlob(attachments[i].id);
                const sf = new FormData();
                sf.append("document", blob, attachments[i].name);
                await fetch(`${BUNDLE_ADD_URL}?conversation_id=${currentConversationId}`, {
                    method:  "POST",
                    headers: { Authorization: `Bearer ${token}` },
                    body:    sf
                });
            }
        }

        // 3. Call Welcome API
        setLoadingText("Analysing documents…");
        const welcomeData = await callWelcomeAPI(token);

        // 4. Show chat
        showView("chat");
        if (welcomeData.message) appendMessage("ai", formatAIResponse(welcomeData.message));
        if (welcomeData.tags && welcomeData.tags.length > 0) renderSuggestions(welcomeData.tags);

    } catch (err) {
        console.error("handleUpload error:", err);
        showError(err.message);
    }
}

// ── Chat ──────────────────────────────────────────────────────────
async function sendChatMessage() {
    const input = document.getElementById("user-input");
    const text  = input.value.trim();
    if (!text) return;

    input.value = "";
    input.style.height = "auto";
    hideSuggestions();
    appendMessage("user", text);

    const typingId = appendTypingIndicator();
    document.getElementById("btn-send").disabled = true;

    try {
        const token = await getAuthToken();
        const resp  = await fetch(ASK_URL, {
            method:  "POST",
            headers: { "Content-Type": "application/json", Authorization: `Bearer ${token}` },
            body:    JSON.stringify({
                conversationId: currentConversationId,
                text:           text,
                isMobile:       false
            })
        });

        removeTypingIndicator(typingId);
        const data   = await resp.json();
        const answer = data.answer || data.response || data.message || "No response received.";
        appendMessage("ai", formatAIResponse(answer));
        if (data.tags && data.tags.length > 0) renderSuggestions(data.tags);

    } catch (err) {
        removeTypingIndicator(typingId);
        appendMessage("ai", "Error: " + err.message);
    } finally {
        document.getElementById("btn-send").disabled = false;
    }
}

// ── API Calls ─────────────────────────────────────────────────────
async function callWelcomeAPI(token) {
    try {
        const resp = await fetch(WELCOME_URL, {
            method:  "POST",
            headers: { "Content-Type": "application/json", Authorization: `Bearer ${token}` },
            body:    JSON.stringify({
                conversationId: currentConversationId,
                documentId:     currentDocumentId
            })
        });
        if (!resp.ok) return { message: "Analysis complete. How can I help?", tags: [] };
        const data = await resp.json();
        return {
            message: data.message || data.answer || data.response || "How can I help?",
            tags:    data.tags || []
        };
    } catch (err) {
        console.warn("Welcome API error:", err);
        return { message: "How can I help?", tags: [] };
    }
}

function getAttachmentBlob(attachmentId) {
    return new Promise((resolve, reject) => {
        Office.context.mailbox.item.getAttachmentContentAsync(attachmentId, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                const binary = atob(result.value.content);
                const bytes  = new Uint8Array(binary.length);
                for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
                resolve(new Blob([bytes]));
            } else {
                reject(new Error(result.error.message));
            }
        });
    });
}

// ── UI Rendering ──────────────────────────────────────────────────
function appendMessage(role, html) {
    const hist = document.getElementById("chat-history");
    const div  = document.createElement("div");
    div.className = role === "user" ? "msg-row msg-user" : "msg-row msg-ai";
    div.innerHTML = `<div class="bubble">${html}</div>`;
    hist.appendChild(div);
    document.getElementById("chat-window").scrollTop = hist.scrollHeight;
}

function appendTypingIndicator() {
    const id   = "typing-" + Date.now();
    const hist = document.getElementById("chat-history");
    const div  = document.createElement("div");
    div.id        = id;
    div.className = "msg-row msg-ai";
    div.innerHTML = `<div class="bubble typing-indicator"><span></span><span></span><span></span></div>`;
    hist.appendChild(div);
    document.getElementById("chat-window").scrollTop = hist.scrollHeight;
    return id;
}

function removeTypingIndicator(id) {
    const el = document.getElementById(id);
    if (el) el.remove();
}

function renderSuggestions(tags) {
    const list = document.getElementById("suggestions-list");
    const bar  = document.getElementById("suggestions-bar");
    list.innerHTML = "";
    tags.forEach(t => {
        const q = typeof t === "string" ? t : (t["next-question"] || "");
        if (!q.trim()) return;
        const btn = document.createElement("button");
        btn.className = "suggestion-btn";
        btn.innerText = q;
        btn.onclick   = () => {
            document.getElementById("user-input").value = q;
            sendChatMessage();
        };
        list.appendChild(btn);
    });
    bar.classList.remove("hidden");
}

function hideSuggestions() {
    document.getElementById("suggestions-bar").classList.add("hidden");
    document.getElementById("suggestions-list").innerHTML = "";
}

function formatAIResponse(text) {
    if (!text) return "";
    return text
        .replace(/\*\*(.*?)\*\*/g, "<b>$1</b>")
        .replace(/__(.*?)__/g,     "<b>$1</b>")
        .replace(/\*(.*?)\*/g,     "<i>$1</i>")
        .replace(/^### (.*?)$/gm,  "<b>$1</b>")
        .replace(/^## (.*?)$/gm,   "<b>$1</b>")
        .replace(/^# (.*?)$/gm,    "<b>$1</b>")
        .replace(/^[-*] (.*?)$/gm, "• $1")
        .replace(/^\d+\. (.*?)$/gm,"• $1")
        .replace(/\n/g,            "<br>");
}

// ── View Management ───────────────────────────────────────────────
function showView(name) {
    document.querySelectorAll(".view").forEach(v => v.classList.add("hidden"));
    document.getElementById("view-" + name).classList.remove("hidden");
}

function showError(msg) {
    document.getElementById("error-msg").innerText = msg || "Unknown error.";
    showView("error");
}

function resetToAttachments() {
    currentConversationId = null;
    currentDocumentId     = null;
    document.getElementById("chat-history").innerHTML = "";
    hideSuggestions();
    showView("attachments");
}

function setLoadingText(msg) {
    document.getElementById("loading-text").innerText = msg;
}

// ── Helpers ───────────────────────────────────────────────────────
function formatBytes(b) {
    if (!b)      return "";
    if (b < 1024)    return b + " B";
    if (b < 1048576) return (b / 1024).toFixed(1) + " KB";
    return (b / 1048576).toFixed(1) + " MB";
}

function getIconForName(name) {
    const ext   = (name || "").split(".").pop().toLowerCase();
    const icons = {
        pdf: "📕", doc: "📘", docx: "📘",
        xls: "📗", xlsx: "📗", ppt: "📙", pptx: "📙",
        png: "🖼️", jpg: "🖼️", jpeg: "🖼️", gif: "🖼️",
        zip: "📦", rar: "📦", mp4: "🎬", mp3: "🎵"
    };
    return icons[ext] || "📄";
}