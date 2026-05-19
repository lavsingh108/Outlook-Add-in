// ── Proxy base URL ─────────────────────────────────────────────────────────
const PROXY_BASE     = "https://headphone-crust-stipulate.ngrok-free.dev";

const AUTH_URL       = `${PROXY_BASE}/v1/authenticate`;
const UPLOAD_URL     = `${PROXY_BASE}/v1/document/upload`;
const BUNDLE_ADD_URL = `${PROXY_BASE}/v1/document/bundle/add`;
const WELCOME_URL    = `${PROXY_BASE}/v1/conversation/ask/welcome`;
const ASK_URL        = `${PROXY_BASE}/v1/conversation/ask/question`;

// ── Share API — update path/body once spec is provided ────────────────────
// POST  SHARE_URL
// Body (JSON): { doc_id, sender_email, recipients: ["a@b.com", …] }
// Expected:    { share_url: "https://…" }
const SHARE_URL = `${PROXY_BASE}/v1/document/share`;

// ── MSAL Config ────────────────────────────────────────────────────────────
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

// ── Shared state ───────────────────────────────────────────────────────────
let _cachedSmartBlueToken = null;
let currentConversationId = null;
let currentDocumentId     = null;

// ══════════════════════════════════════════════════════════════════════════
// ENTRY POINT
// ══════════════════════════════════════════════════════════════════════════
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) init();
});

function init() {
    // Compose mode exposes setAsync on subject/body; read mode does not.
    const item = Office.context.mailbox.item;
    const isCompose = typeof item.subject?.setAsync === "function"
                   || typeof item.body?.setAsync    === "function";

    if (isCompose) {
        initCompose();
    } else {
        initRead();
    }
}

// ══════════════════════════════════════════════════════════════════════════
// READ MODE
// ══════════════════════════════════════════════════════════════════════════
function initRead() {
    loadAttachments();
    document.getElementById("btn-upload-bundle").onclick = handleBundleUpload;
    document.getElementById("btn-send").onclick          = sendChatMessage;
    document.getElementById("btn-back").onclick          = switchToAttachments;
    document.getElementById("chk-bulk").onchange         = onToggleMode;
    document.getElementById("user-input").addEventListener("keydown", (e) => {
        if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); sendChatMessage(); }
    });
}

function isBulkMode() { return document.getElementById("chk-bulk").checked; }

function onToggleMode() {
    const bulk = isBulkMode();
    document.getElementById("lbl-bundle").classList.toggle("active", bulk);
    document.getElementById("lbl-individual").classList.toggle("active", !bulk);
    document.getElementById("bundle-footer").classList.toggle("hidden", !bulk);
    loadAttachments();
}

function loadAttachments() {
    // const attachments = Office.context.mailbox.item.attachments;
    const attachments = Office.context.mailbox.item.getAttachmentsAsync();
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

    container.querySelectorAll("input[name='primaryIndex']").forEach(radio => {
        radio.addEventListener("change", () => updateBundleSelection(container));
    });
}

function updateBundleSelection(container) {
    const primaryVal = container.querySelector("input[name='primaryIndex']:checked")?.value;
    container.querySelectorAll(".att-item").forEach(item => {
        const idx       = item.dataset.index;
        const isPrimary = idx === primaryVal;
        const secChk    = item.querySelector("input[name='secondaryIndex']");
        item.classList.toggle("is-primary", isPrimary);
        if (isPrimary) {
            secChk.checked  = false;
            secChk.disabled = true;
        } else {
            secChk.disabled = false;
            if (!secChk.dataset.userUnchecked) secChk.checked = true;
        }
    });
}

document.addEventListener("change", (e) => {
    if (e.target.name === "secondaryIndex") {
        e.target.dataset.userUnchecked = e.target.checked ? "" : "1";
    }
});

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

async function handleBundleUpload() {
    const attachments  = Office.context.mailbox.item.attachments;
    const primaryRadio = document.querySelector("input[name='primaryIndex']:checked");
    if (!primaryRadio) { showStatus("Please select a primary document."); return; }

    const primaryIndex     = parseInt(primaryRadio.value);
    const primaryAtt       = attachments[primaryIndex];
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

// ══════════════════════════════════════════════════════════════════════════
// COMPOSE MODE
// ══════════════════════════════════════════════════════════════════════════

let _composeAttachments = [];   // Office attachment objects from the email
let _composeRecipients  = [];   // flat array of email strings (To + CC)
let _senderEmail        = "";

function initCompose() {
    document.querySelector(".header-title").textContent = "Share Document";

    document.getElementById("view-compose").classList.remove("hidden");
    document.getElementById("view-attachments").classList.add("hidden");

    document.getElementById("btn-refresh").classList.remove("hidden");
    document.getElementById("btn-refresh").onclick = () => loadComposeData(true);

    document.getElementById("btn-compose-upload").onclick = handleComposeUpload;
    document.getElementById("btn-copy-link").onclick      = copyResultLink;

    loadComposeData(false);
}

// ── Read recipients + attachments from the live compose email ─────────────
function loadComposeData(isRefresh) {
    if (isRefresh) {
        document.getElementById("btn-refresh").classList.add("spinning");
        document.getElementById("compose-result").classList.add("hidden");
        showComposeStatus("");
    }

    try {
        _senderEmail = Office.context.mailbox.userProfile.emailAddress || "";
    } catch (e) { _senderEmail = ""; }

    const item = Office.context.mailbox.item;

    Promise.all([
        new Promise(res => item.to.getAsync(r =>
            res(r.status === Office.AsyncResultStatus.Succeeded ? r.value : []))),
        new Promise(res => item.cc.getAsync(r =>
            res(r.status === Office.AsyncResultStatus.Succeeded ? r.value : [])))
    ]).then(([toList, ccList]) => {
        // Deduplicated flat list of emails for the share API payload
        const seen = new Set();
        _composeRecipients = [...toList, ...ccList]
            .map(r => (r.emailAddress || "").toLowerCase().trim())
            .filter(e => { if (!e || seen.has(e)) return false; seen.add(e); return true; });

        renderComposeRecipients(toList, ccList);

        _composeAttachments = item.attachments || [];
        renderComposeAttachments(_composeAttachments);

        document.getElementById("btn-refresh").classList.remove("spinning");
        validateComposeReady();
    });
}

// ── Render recipient chips (To / CC rows) ─────────────────────────────────
function renderComposeRecipients(toList, ccList) {
    const area  = document.getElementById("compose-recipients");
    const badge = document.getElementById("recipients-count");
    const total = toList.length + ccList.length;

    badge.textContent = total || "";

    if (total === 0) {
        area.innerHTML = `<div class="compose-empty">
            No recipients yet. Add To / CC addresses then click ↻ Refresh.
        </div>`;
        return;
    }

    area.innerHTML = "";

    const buildRow = (label, list) => {
        if (!list.length) return;
        const row = document.createElement("div");
        row.className = "recipient-row";

        const lbl = document.createElement("span");
        lbl.className = "recipient-row-label";
        lbl.textContent = label;
        row.appendChild(lbl);

        const chips = document.createElement("div");
        chips.className = "recipient-chips";
        list.forEach(r => {
            const chip = document.createElement("span");
            chip.className = "recipient-chip";
            chip.textContent = r.emailAddress || r.displayName || "";
            chip.title = r.emailAddress || "";
            chips.appendChild(chip);
        });
        row.appendChild(chips);
        area.appendChild(row);
    };

    buildRow("To:", toList);
    buildRow("CC:", ccList);
}

// ── Render attachment radio cards ─────────────────────────────────────────
function renderComposeAttachments(attachments) {
    const list  = document.getElementById("compose-attachments");
    const badge = document.getElementById("attachments-count");

    badge.textContent = attachments.length || "";

    if (!attachments.length) {
        list.innerHTML = `<div class="compose-empty">
            No attachments yet. Attach a document to the email then click ↻ Refresh.
        </div>`;
        document.getElementById("btn-compose-upload").disabled = true;
        return;
    }

    list.innerHTML = "";

    attachments.forEach((att, index) => {
        const div = document.createElement("div");
        div.className = "compose-att-item" + (index === 0 ? " selected" : "");
        div.innerHTML = `
            <input type="radio" name="composeDoc" id="cdoc-${index}"
                   value="${index}" ${index === 0 ? "checked" : ""}/>
            <label for="cdoc-${index}" class="compose-att-label">
                <svg class="compose-att-icon" viewBox="0 0 24 24" fill="none"
                     stroke="currentColor" stroke-width="2"
                     stroke-linecap="round" stroke-linejoin="round">
                    <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                    <polyline points="14 2 14 8 20 8"/>
                </svg>
                <div class="compose-att-meta">
                    <div class="compose-att-name" title="${att.name}">${att.name}</div>
                    <div class="compose-att-size">${formatBytes(att.size)}</div>
                </div>
            </label>`;
        list.appendChild(div);

        div.querySelector("input[type=radio]").addEventListener("change", () => {
            list.querySelectorAll(".compose-att-item").forEach(el => el.classList.remove("selected"));
            div.classList.add("selected");
            validateComposeReady();
        });
    });
}

function validateComposeReady() {
    const hasAtt = _composeAttachments.length > 0
                && !!document.querySelector("input[name='composeDoc']:checked");
    document.getElementById("btn-compose-upload").disabled = !hasAtt;
}

// ── Upload → Share API → Insert into body ─────────────────────────────────
async function handleComposeUpload() {
    const radioEl = document.querySelector("input[name='composeDoc']:checked");
    if (!radioEl) return;

    const att       = _composeAttachments[parseInt(radioEl.value)];
    const uploadBtn = document.getElementById("btn-compose-upload");

    uploadBtn.disabled = true;
    document.getElementById("compose-result").classList.add("hidden");

    try {
        // Step 1 — upload
        showComposeStatus("Signing in…");
        const token = await getAuthToken();

        showComposeStatus("Uploading document…");
        const blob = await getAttachmentBlob(att.id, att.name);
        const form = new FormData();
        form.append("document", blob, att.name);

        const uploadResp = await fetch(UPLOAD_URL, {
            method:  "POST",
            headers: { Authorization: "Bearer " + token },
            body:    form,
        });
        if (!uploadResp.ok)
            throw new Error("Upload failed (" + uploadResp.status + "): " + await uploadResp.text());

        const uploadData = await uploadResp.json();
        const docId = uploadData.doc_id || uploadData.documentId || uploadData.id || "";
        if (!docId) throw new Error("No document ID returned by upload.");

        // Step 2 — share API
        showComposeStatus("Creating share link…");
        const shareLink = await callShareApi(token, docId);

        // Step 3 — insert into body
        showComposeStatus("Inserting link into email…");
        await insertShareLinkIntoBody(shareLink, att.name);

        showComposeStatus("");
        renderComposeResult(shareLink);

    } catch (err) {
        console.error("Compose upload error:", err);
        showComposeStatus("Error: " + err.message);
        _cachedSmartBlueToken = null;
    } finally {
        uploadBtn.disabled = false;
    }
}

// ── Share API ──────────────────────────────────────────────────────────────
// ⚠  Update endpoint path and request body once the API spec is provided.
// Current shape: POST /v1/document/share
//   { doc_id, sender_email, recipients: [...] }  →  { share_url }
async function callShareApi(token, docId) {
    const resp = await fetch(SHARE_URL, {
        method:  "POST",
        headers: {
            "Content-Type": "application/json",
            Authorization:  "Bearer " + token,
        },
        body: JSON.stringify({
            doc_id:       docId,
            sender_email: _senderEmail,
            recipients:   _composeRecipients,
        }),
    });

    if (!resp.ok)
        throw new Error("Share API failed (" + resp.status + "): " + await resp.text());

    const data = await resp.json();
    const url  = data.share_url || data.shareUrl || data.url || "";
    if (!url) throw new Error("Share API returned no URL.");
    return url;
}

// ── Auto-insert share link into the email body ─────────────────────────────
function insertShareLinkIntoBody(link, filename) {
    return new Promise((resolve) => {
        const html = `<p style="font-family:sans-serif;margin:8px 0;">`
                   + `<a href="${link}" target="_blank" `
                   + `style="color:#0D47A1;font-weight:600;text-decoration:none;">`
                   + `📄 ${filename} — View on SmartBlue</a></p>`;

        Office.context.mailbox.item.body.setSelectedDataAsync(
            html,
            { coercionType: Office.CoercionType.Html },
            (result) => {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    // Plain-text fallback
                    Office.context.mailbox.item.body.setSelectedDataAsync(
                        `\n${filename}: ${link}\n`,
                        { coercionType: Office.CoercionType.Text },
                        () => resolve()
                    );
                } else {
                    resolve();
                }
            }
        );
    });
}

function renderComposeResult(link) {
    document.getElementById("result-link-text").textContent = link;
    document.getElementById("compose-result").classList.remove("hidden");
    document.getElementById("compose-result").scrollIntoView({ behavior: "smooth" });
}

function copyResultLink() {
    const link = document.getElementById("result-link-text").textContent;
    const btn  = document.getElementById("btn-copy-link");
    const done = () => {
        btn.classList.add("copied");
        setTimeout(() => btn.classList.remove("copied"), 1600);
    };
    if (navigator.clipboard) {
        navigator.clipboard.writeText(link).then(done).catch(() => fallbackCopy(link, done));
    } else {
        fallbackCopy(link, done);
    }
}

function fallbackCopy(text, cb) {
    const ta = document.createElement("textarea");
    ta.value = text; ta.style.cssText = "position:fixed;opacity:0";
    document.body.appendChild(ta); ta.select();
    try { document.execCommand("copy"); cb(); } catch (_) {}
    document.body.removeChild(ta);
}

function showComposeStatus(msg) {
    document.getElementById("compose-status").innerText = msg;
}

// ══════════════════════════════════════════════════════════════════════════
// AUTH
// ══════════════════════════════════════════════════════════════════════════
async function getAuthToken() {
    if (_cachedSmartBlueToken) return _cachedSmartBlueToken;

    const msalInstance = getMsal();
    let idToken = null;

    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
        try {
            const silent = await msalInstance.acquireTokenSilent({ scopes: SCOPES, account: accounts[0] });
            idToken = silent.idToken;
        } catch (e) { console.warn("Silent token failed:", e.message); }
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
        method:  "POST",
        headers: { "Content-Type": "application/json" },
        body:    JSON.stringify({ idToken }),
    });
    if (!authResp.ok) throw new Error("Auth failed (" + authResp.status + "): " + await authResp.text());

    const { token } = await authResp.json();
    if (!token) throw new Error("No token returned from auth proxy.");
    _cachedSmartBlueToken = token;
    return token;
}

// ══════════════════════════════════════════════════════════════════════════
// UPLOAD HELPERS  (shared by read + compose)
// ══════════════════════════════════════════════════════════════════════════
async function uploadPrimary(att, token) {
    const blob = await getAttachmentBlob(att.id, att.name);
    const form = new FormData();
    form.append("document", blob, att.name);

    const resp = await fetch(UPLOAD_URL, {
        method:  "POST",
        headers: { Authorization: "Bearer " + token },
        body:    form,
    });
    if (!resp.ok) throw new Error("Upload failed (" + resp.status + "): " + await resp.text());

    const data           = await resp.json();
    const conversationId = data.conversation_id || data.conversationId;
    const documentId     = data.doc_id || data.documentId || data.id || null;

    if (!conversationId) throw new Error("No conversation_id returned by upload.");
    return { conversationId, documentId };
}

async function uploadSupporting(att, conversationId, token) {
    const blob = await getAttachmentBlob(att.id, att.name);
    const form = new FormData();
    form.append("document", blob, att.name);

    const resp = await fetch(`${BUNDLE_ADD_URL}?conversation_id=${encodeURIComponent(conversationId)}`, {
        method:  "POST",
        headers: { Authorization: "Bearer " + token },
        body:    form,
    });
    if (!resp.ok) console.warn("Supporting upload failed for:", att.name, await resp.text());
}

function getAttachmentBlob(attachmentId, filename) {
    return new Promise((resolve, reject) => {
        Office.context.mailbox.item.getAttachmentContentAsync(attachmentId, (result) => { // getAttachmentsAsync is for compose mode; getAttachmentContentAsync works in both modes
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

function getMimeType(filename) {
    const ext = (filename || "").split(".").pop().toLowerCase();
    const MAP = {
        pdf: "application/pdf", doc: "application/msword",
        docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        xls: "application/vnd.ms-excel",
        xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        csv: "text/csv", ppt: "application/vnd.ms-powerpoint",
        pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        txt: "text/plain", rtf: "application/rtf",
        png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg",
        gif: "image/gif", webp: "image/webp", zip: "application/zip",
    };
    return MAP[ext] || "application/octet-stream";
}

// ══════════════════════════════════════════════════════════════════════════
// READ MODE — CHAT
// ══════════════════════════════════════════════════════════════════════════
async function enterChat(conversationId, documentId, token) {
    currentConversationId = conversationId;
    currentDocumentId     = documentId;
    switchToChat();
    showStatus("Loading…");
    showTypingIndicator();

    try {
        const resp = await fetch(WELCOME_URL, {
            method:  "POST",
            headers: { "Content-Type": "application/json", Authorization: "Bearer " + token },
            body:    JSON.stringify({ conversationId, documentId }),
        });
        hideTypingIndicator();

        if (resp.ok) {
            const data = await resp.json();
            appendMessage("ai", data.answer || data.response || data.message || "How can I help you today?");
            const tags = Array.isArray(data.tags) ? data.tags : [];
            if (tags.length) renderSuggestions(tags);
        } else {
            appendMessage("ai", "Document uploaded. How can I help you?");
        }
    } catch (err) {
        hideTypingIndicator();
        appendMessage("ai", "Document uploaded. How can I help you?");
    }
    showStatus("");
}

function renderSuggestions(tags) {
    const box = document.getElementById("suggestions");
    box.innerHTML = "";
    tags.forEach(tag => {
        const q = typeof tag === "string" ? tag : (tag["next-question"] || tag.question || "");
        if (!q.trim()) return;
        const chip = document.createElement("button");
        chip.className = "chip";
        chip.textContent = q;
        chip.onclick = () => { hideSuggestions(); document.getElementById("user-input").value = q; sendChatMessage(); };
        box.appendChild(chip);
    });
    box.classList.remove("hidden");
}

function hideSuggestions() {
    const box = document.getElementById("suggestions");
    box.classList.add("hidden");
    box.innerHTML = "";
}

async function sendChatMessage() {
    const input = document.getElementById("user-input");
    const text  = input.value.trim();
    if (!text) return;

    hideSuggestions();
    appendMessage("user", text);
    input.value = "";
    document.getElementById("btn-send").disabled = true;
    showTypingIndicator();

    try {
        const token = await getAuthToken();
        const resp  = await fetch(ASK_URL, {
            method:  "POST",
            headers: { "Content-Type": "application/json", Authorization: "Bearer " + token },
            body:    JSON.stringify({ conversationId: currentConversationId, text, isMobile: false }),
        });
        hideTypingIndicator();
        if (!resp.ok) throw new Error("Ask failed (" + resp.status + "): " + await resp.text());
        const data = await resp.json();
        appendMessage("ai", data.answer || data.response || "No response received.");
        const tags = Array.isArray(data.tags) ? data.tags : [];
        if (tags.length) renderSuggestions(tags);
    } catch (err) {
        hideTypingIndicator();
        appendMessage("ai", "Error: " + err.message);
        _cachedSmartBlueToken = null;
    } finally {
        document.getElementById("btn-send").disabled = false;
    }
}

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

function showTypingIndicator() {
    const hist = document.getElementById("chat-history");
    if (hist.querySelector(".msg-typing")) return;
    const div = document.createElement("div");
    div.className = "msg-typing"; div.id = "typing-indicator";
    div.innerHTML = `<span class="typing-dot"></span><span class="typing-dot"></span><span class="typing-dot"></span>`;
    hist.appendChild(div); hist.scrollTop = hist.scrollHeight;
}

function hideTypingIndicator() {
    const el = document.getElementById("typing-indicator");
    if (el) el.remove();
}

function appendMessage(role, text) {
    const hist = document.getElementById("chat-history");
    const div  = document.createElement("div");
    if (role === "user") {
        div.className = "msg-user";
        const p = document.createElement("p"); p.textContent = text; div.appendChild(p);
    } else {
        div.className = "msg-ai"; div.innerHTML = formatResponse(text);
    }
    hist.appendChild(div); hist.scrollTop = hist.scrollHeight;
}

function switchToChat() {
    document.getElementById("view-attachments").classList.add("hidden");
    document.getElementById("view-chat").classList.remove("hidden");
    document.getElementById("btn-back").classList.remove("hidden");
    document.getElementById("chat-history").innerHTML = "";
    hideSuggestions(); showStatus("");
}

function switchToAttachments() {
    document.getElementById("view-chat").classList.add("hidden");
    document.getElementById("view-attachments").classList.remove("hidden");
    document.getElementById("btn-back").classList.add("hidden");
    document.getElementById("chat-history").innerHTML = "";
    hideSuggestions();
    currentConversationId = null; currentDocumentId = null;
    document.getElementById("btn-upload-bundle").disabled = false;
    loadAttachments(); showStatus("");
}

function showStatus(msg) { document.getElementById("status-msg").innerText = msg; }

function formatBytes(bytes) {
    if (!bytes) return "";
    if (bytes < 1024)    return bytes + " B";
    if (bytes < 1048576) return (bytes / 1024).toFixed(1) + " KB";
    return (bytes / 1048576).toFixed(1) + " MB";
}