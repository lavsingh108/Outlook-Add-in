// ── Proxy base URL ────────────────────────────────────────────────
const PROXY_BASE     = "https://headphone-crust-stipulate.ngrok-free.dev";

const AUTH_URL       = `${PROXY_BASE}/v1/authenticate`;
const UPLOAD_URL     = `${PROXY_BASE}/v1/document/upload`;
const BUNDLE_ADD_URL = `${PROXY_BASE}/v1/document/bundle/add`;
const WELCOME_URL    = `${PROXY_BASE}/v1/conversation/ask/welcome`;
const ASK_URL        = `${PROXY_BASE}/v1/conversation/ask/question`;

// ── MSAL Config ───────────────────────────────────────────────────
const AZURE_CLIENT_ID = "c49037f2-0565-4a5c-8b17-f9b8b3ee35c7";
const AZURE_TENANT_ID = "f895e126-dbc8-41bb-b00b-5cd2172346f9";
const SCOPES = ["openid", "profile", "email", "User.Read"];

const msalConfig = {
    auth: {
        clientId: AZURE_CLIENT_ID,
        authority: "https://login.microsoftonline.com/" + AZURE_TENANT_ID,
        redirectUri: window.location.href.split("?")[0]
    },
    cache: { cacheLocation: "sessionStorage", storeAuthStateInCookie: false }
};

let _msal = null;
function getMsal() {
    if (!_msal) _msal = new msal.PublicClientApplication(msalConfig);
    return _msal;
}

// ── State ─────────────────────────────────────────────────────────
let _cachedSmartBlueToken = null;
let currentConversationId = null;
let currentDocumentId     = null;

// ── Entry Point ───────────────────────────────────────────────────
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) init();
});

function init() {
    loadAttachments();
    document.getElementById("btn-upload-bundle").onclick = handleBundleUpload;
    document.getElementById("btn-send").onclick          = sendChatMessage;
    document.getElementById("btn-back").onclick          = switchToAttachments;
    document.getElementById("chk-bulk").onchange         = onToggleMode;
    document.getElementById("user-input").addEventListener("keydown", (e) => {
        if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); sendChatMessage(); }
    });
}

// ── Mode toggle ───────────────────────────────────────────────────
function isBulkMode() { return document.getElementById("chk-bulk").checked; }

function onToggleMode() {
    const bulk = isBulkMode();
    document.getElementById("lbl-bundle").classList.toggle("active", bulk);
    document.getElementById("lbl-individual").classList.toggle("active", !bulk);
    document.getElementById("bundle-footer").classList.toggle("hidden", !bulk);
    loadAttachments();
}

// ── Load & render attachments ─────────────────────────────────────
function loadAttachments() {
    const attachments = Office.context.mailbox.item.attachments;
    const listDiv     = document.getElementById("attachment-list");

    if (!attachments || attachments.length === 0) {
        listDiv.innerHTML = "<p style='color:#888;font-size:13px;padding:4px 0'>No attachments found.</p>";
        document.getElementById("btn-upload-bundle").disabled = true;
        return;
    }

    document.getElementById("btn-upload-bundle").disabled = false;
    listDiv.innerHTML = "";

    if (isBulkMode()) renderBundleList(attachments, listDiv);
    else              renderIndividualList(attachments, listDiv);
}

// Bundle mode: radio (primary) + checkbox (secondary)
function renderBundleList(attachments, container) {
    attachments.forEach((att, index) => {
        const isPrimary = index === 0;
        const div = document.createElement("div");
        div.className = "att-item" + (isPrimary ? " is-primary" : "");
        div.dataset.index = index;

        div.innerHTML = `
            <div class="att-bundle-row">
                <div class="att-radio-col">
                    <input type="radio" name="primaryIndex" value="${index}"
                           id="radio-${index}" ${isPrimary ? "checked" : ""}/>
                    <label class="radio-label" for="radio-${index}">Primary</label>
                </div>
                <div class="att-info">
                    <div class="att-name" title="${att.name}">${att.name}</div>
                    <div class="att-meta">${formatBytes(att.size)}</div>
                </div>
                <div class="att-secondary-col">
                    <input type="checkbox" name="secondaryIndex" value="${index}"
                           id="chk-sec-${index}" ${isPrimary ? "" : "checked"}
                           ${isPrimary ? "disabled" : ""}/>
                    <label class="sec-label" for="chk-sec-${index}">Include</label>
                </div>
            </div>`;

        container.appendChild(div);
    });

    // When primary radio changes: highlight the new primary row,
    // uncheck + disable secondary on it, re-enable all others
    container.querySelectorAll("input[name='primaryIndex']").forEach(radio => {
        radio.addEventListener("change", () => updateBundleSelection(container));
    });
}

function updateBundleSelection(container) {
    const primaryVal = container.querySelector("input[name='primaryIndex']:checked")?.value;
    container.querySelectorAll(".att-item").forEach(item => {
        const idx      = item.dataset.index;
        const isPrimary = idx === primaryVal;
        const secChk   = item.querySelector("input[name='secondaryIndex']");

        item.classList.toggle("is-primary", isPrimary);

        if (isPrimary) {
            secChk.checked  = false;
            secChk.disabled = true;
        } else {
            secChk.disabled = false;
            // Re-check it if it was forcibly unchecked when it became primary before
            if (!secChk.dataset.userUnchecked) secChk.checked = true;
        }
    });
}

// Track manual unchecks so we don't re-check them on radio change
document.addEventListener("change", (e) => {
    if (e.target.name === "secondaryIndex") {
        e.target.dataset.userUnchecked = e.target.checked ? "" : "1";
    }
});

// Individual mode: each row has its own Upload button
function renderIndividualList(attachments, container) {
    attachments.forEach((att, index) => {
        const div = document.createElement("div");
        div.className = "att-item";
        div.innerHTML = `
            <div class="att-individual-row">
                <div class="att-info">
                    <div class="att-name" title="${att.name}">${att.name}</div>
                    <div class="att-meta">${formatBytes(att.size)}</div>
                </div>
                <button class="btn-upload-single" data-index="${index}">Upload</button>
            </div>`;
        container.appendChild(div);
    });

    container.querySelectorAll(".btn-upload-single").forEach(btn => {
        btn.onclick = () => handleSingleUpload(parseInt(btn.dataset.index));
    });
}

// ── Auth ──────────────────────────────────────────────────────────
async function getAuthToken() {
    if (_cachedSmartBlueToken) return _cachedSmartBlueToken;

    const msalInstance = getMsal();
    let idToken = null;

    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
        try {
            const silent = await msalInstance.acquireTokenSilent({ scopes: SCOPES, account: accounts[0] });
            idToken = silent.idToken;
        } catch (e) { console.warn("Silent failed:", e.message); }
    }

    if (!idToken) {
        try {
            const popup = await msalInstance.loginPopup({ scopes: SCOPES, prompt: "select_account" });
            idToken = popup.idToken;
        } catch (e) {
            throw new Error("Sign-in failed: " + (e.message || e.errorCode));
        }
    }

    const authResp = await fetch(AUTH_URL, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ idToken }),
    });
    if (!authResp.ok) throw new Error("Auth failed (" + authResp.status + "): " + await authResp.text());

    const { token } = await authResp.json();
    if (!token) throw new Error("No token returned from auth proxy");

    _cachedSmartBlueToken = token;
    return token;
}

// ── Bundle upload ─────────────────────────────────────────────────
async function handleBundleUpload() {
    const attachments  = Office.context.mailbox.item.attachments;
    const primaryRadio = document.querySelector("input[name='primaryIndex']:checked");
    if (!primaryRadio) { showStatus("Please select a primary document."); return; }

    const primaryIndex = parseInt(primaryRadio.value);
    const primaryAtt   = attachments[primaryIndex];

    // Collect checked secondary indices (excluding primary)
    const secondaryIndices = Array.from(
        document.querySelectorAll("input[name='secondaryIndex']:checked")
    ).map(c => parseInt(c.value)).filter(i => i !== primaryIndex);

    showStatus("Signing in…");
    document.getElementById("btn-upload-bundle").disabled = true;

    try {
        const token = await getAuthToken();

        showStatus("Uploading primary document…");
        const { conversationId, documentId } = await uploadPrimary(primaryAtt, token);

        if (secondaryIndices.length > 0) {
            showStatus(`Uploading ${secondaryIndices.length} secondary document(s)…`);
            for (const idx of secondaryIndices) {
                await uploadSupporting(attachments[idx], conversationId, token);
            }
        }

        await enterChat(conversationId, documentId, token);

    } catch (err) {
        console.error("Bundle upload error:", err);
        showStatus("Error: " + err.message);
        _cachedSmartBlueToken = null;
        document.getElementById("btn-upload-bundle").disabled = false;
    }
}

// ── Single upload ─────────────────────────────────────────────────
async function handleSingleUpload(index) {
    const att = Office.context.mailbox.item.attachments[index];
    document.querySelectorAll(".btn-upload-single").forEach(b => b.disabled = true);
    showStatus("Uploading " + att.name + "…");

    try {
        const token = await getAuthToken();
        const { conversationId, documentId } = await uploadPrimary(att, token);
        await enterChat(conversationId, documentId, token);
    } catch (err) {
        console.error("Single upload error:", err);
        showStatus("Error: " + err.message);
        _cachedSmartBlueToken = null;
        document.querySelectorAll(".btn-upload-single").forEach(b => b.disabled = false);
    }
}

// ── Upload helpers ────────────────────────────────────────────────
async function uploadPrimary(att, token) {
    const blob = await getAttachmentBlob(att.id, att.name);
    const form = new FormData();
    form.append("document", blob, att.name);

    const resp = await fetch(UPLOAD_URL, {
        method: "POST",
        headers: { Authorization: "Bearer " + token },
        body: form,
    });
    if (!resp.ok) throw new Error("Upload failed (" + resp.status + "): " + await resp.text());

    const data         = await resp.json();
    const conversationId = data.conversation_id || data.conversationId;
    const documentId     = data.doc_id || data.documentId || data.id || null;

    if (!conversationId) throw new Error("No conversation_id returned by upload");
    console.log("Uploaded primary. conversation_id:", conversationId, "doc_id:", documentId);
    return { conversationId, documentId };
}

async function uploadSupporting(att, conversationId, token) {
    const blob = await getAttachmentBlob(att.id, att.name);
    const form = new FormData();
    form.append("document", blob, att.name);

    const resp = await fetch(`${BUNDLE_ADD_URL}?conversation_id=${encodeURIComponent(conversationId)}`, {
        method: "POST",
        headers: { Authorization: "Bearer " + token },
        body: form,
    });
    if (!resp.ok) console.warn("Supporting upload failed for:", att.name, await resp.text());
}

// ── Welcome API → enter chat ──────────────────────────────────────
async function enterChat(conversationId, documentId, token) {
    // ── CRITICAL: reset state for the new document ────────────────
    currentConversationId = conversationId;
    currentDocumentId     = documentId;

    switchToChat();  // clears chat history, hides suggestions
    showStatus("Loading…");

    try {
        const resp = await fetch(WELCOME_URL, {
            method: "POST",
            headers: { "Content-Type": "application/json", Authorization: "Bearer " + token },
            body: JSON.stringify({ conversationId, documentId }),
        });

        if (resp.ok) {
            const data = await resp.json();
            const msg  = data.answer || data.response || data.message || "How can I help you today?";
            const tags = Array.isArray(data.tags) ? data.tags : [];
            appendMessage("ai", msg);
            if (tags.length > 0) renderSuggestions(tags);
        } else {
            appendMessage("ai", "Document uploaded. How can I help you?");
        }
    } catch (err) {
        console.warn("Welcome API error:", err.message);
        appendMessage("ai", "Document uploaded. How can I help you?");
    }

    showStatus("");
}

// ── Suggestion chips ──────────────────────────────────────────────
function renderSuggestions(tags) {
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
            sendChatMessage();
        };
        box.appendChild(chip);
    });
    box.classList.remove("hidden");
}

function hideSuggestions() {
    const box = document.getElementById("suggestions");
    box.classList.add("hidden");
    box.innerHTML = "";
}

// ── Chat ──────────────────────────────────────────────────────────
async function sendChatMessage() {
    const input = document.getElementById("user-input");
    const text  = input.value.trim();
    if (!text) return;

    hideSuggestions();
    appendMessage("user", text);
    input.value = "";
    document.getElementById("btn-send").disabled = true;

    try {
        const token = await getAuthToken();
        const resp  = await fetch(ASK_URL, {
            method: "POST",
            headers: { "Content-Type": "application/json", Authorization: "Bearer " + token },
            body: JSON.stringify({ conversationId: currentConversationId, text, isMobile: false }),
        });

        if (!resp.ok) throw new Error("Ask failed (" + resp.status + "): " + await resp.text());

        const data = await resp.json();
        appendMessage("ai", data.answer || data.response || "No response received.");

        const tags = Array.isArray(data.tags) ? data.tags : [];
        if (tags.length > 0) renderSuggestions(tags);

    } catch (err) {
        console.error("Chat error:", err);
        appendMessage("ai", "Error: " + err.message);
        _cachedSmartBlueToken = null;
    } finally {
        document.getElementById("btn-send").disabled = false;
    }
}

// ── MIME type ─────────────────────────────────────────────────────
function getMimeType(filename) {
    const ext = (filename || "").split(".").pop().toLowerCase();
    const MAP = {
        pdf: "application/pdf",
        doc: "application/msword",
        docx:"application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        xls: "application/vnd.ms-excel",
        xlsx:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        csv: "text/csv",
        ppt: "application/vnd.ms-powerpoint",
        pptx:"application/vnd.openxmlformats-officedocument.presentationml.presentation",
        txt: "text/plain", rtf: "application/rtf",
        png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg",
        gif: "image/gif", webp: "image/webp", zip: "application/zip",
    };
    return MAP[ext] || "application/octet-stream";
}

function getAttachmentBlob(attachmentId, filename) {
    return new Promise((resolve, reject) => {
        Office.context.mailbox.item.getAttachmentContentAsync(attachmentId, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                const binary = atob(result.value.content);
                const bytes  = new Uint8Array(binary.length);
                for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
                resolve(new Blob([bytes], { type: getMimeType(filename) }));
            } else {
                reject(new Error(result.error.message));
            }
        });
    });
}

// ── Response formatter ────────────────────────────────────────────
function formatResponse(raw) {
    let text = raw.replace(
        /<blueEmbed-doc-page>[^:]+:[^:]+:(\d+)<\/blueEmbed-doc-page>/g,
        '<span class="page-ref">pg $1</span>'
    );
    text = text.replace(/\*\*(.+?)\*\*/g, "<strong>$1</strong>");

    const lines = text.split(/\n/);
    let html = "", inList = false;
    for (const rawLine of lines) {
        const line = rawLine.trim();
        if (!line) { if (inList) { html += "</ul>"; inList = false; } continue; }
        if (/^[*●•]\s+/.test(line)) {
            if (!inList) { html += '<ul class="ai-list">'; inList = true; }
            html += "<li>" + line.replace(/^[*●•]\s+/, "") + "</li>";
        } else {
            if (inList) { html += "</ul>"; inList = false; }
            html += "<p>" + line + "</p>";
        }
    }
    if (inList) html += "</ul>";
    return html;
}

// ── UI helpers ────────────────────────────────────────────────────
function appendMessage(role, text) {
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

function switchToChat() {
    document.getElementById("view-attachments").classList.add("hidden");
    document.getElementById("view-chat").classList.remove("hidden");
    document.getElementById("btn-back").classList.remove("hidden");
    // Clear any previous chat session
    document.getElementById("chat-history").innerHTML = "";
    hideSuggestions();
    showStatus("");
}

function switchToAttachments() {
    document.getElementById("view-chat").classList.add("hidden");
    document.getElementById("view-attachments").classList.remove("hidden");
    document.getElementById("btn-back").classList.add("hidden");
    // Clear chat state so the next upload starts fresh
    document.getElementById("chat-history").innerHTML = "";
    hideSuggestions();
    currentConversationId = null;
    currentDocumentId     = null;
    document.getElementById("btn-upload-bundle").disabled = false;
    loadAttachments();  // re-render so radio/checkbox state resets too
    showStatus("");
}

function showStatus(msg) { document.getElementById("status-msg").innerText = msg; }

function formatBytes(bytes) {
    if (!bytes) return "";
    if (bytes < 1024)    return bytes + " B";
    if (bytes < 1048576) return (bytes / 1024).toFixed(1) + " KB";
    return (bytes / 1048576).toFixed(1) + " MB";
}