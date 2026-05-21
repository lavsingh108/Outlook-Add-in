// ── Proxy base URL ─────────────────────────────────────────────────────────
const PROXY_BASE     = "https://headphone-crust-stipulate.ngrok-free.dev";

const AUTH_URL       = `${PROXY_BASE}/v1/authenticate`;
const UPLOAD_URL     = `${PROXY_BASE}/v1/document/upload`;
const SHARE_URL      = `${PROXY_BASE}/v1/document/share`;
const WELCOME_URL    = `${PROXY_BASE}/v1/conversation/ask/welcome`;
const ASK_URL        = `${PROXY_BASE}/v1/conversation/ask/question`;

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
    document.getElementById("header-title-text")?.remove();
    document.querySelector(".header-title").textContent = "View Document";

    // Wire up chat controls
    document.getElementById("btn-send").onclick = sendChatMessage;
    document.getElementById("user-input").addEventListener("keydown", (e) => {
        if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); sendChatMessage(); }
    });

    // Show the read-init loading screen
    document.getElementById("view-read-init").classList.remove("hidden");

    startReadMode();
}

// ── Scan email body for a SmartBlue share URL, then open chat ─────────────
async function startReadMode() {
    setReadInitStatus("Looking for share link…");

    let conversationId = null;
    let docId          = null;

    try {
        ({ conversationId, docId } = await extractShareLinkFromBody());
    } catch (err) {
        showReadInitError("Could not read email body: " + err.message);
        return;
    }

    if (!conversationId) {
        showReadInitError("No SmartBlue share link found in this email.");
        return;
    }

    setReadInitStatus("Authenticating…");

    try {
        const token = await getAuthToken();
        await enterChat(conversationId, docId, token);
    } catch (err) {
        console.error("Read mode init error:", err);
        showReadInitError("Error: " + err.message);
        _cachedSmartBlueToken = null;
    }
}

// ── Read the email body and return the first SmartBlue share link ──────────
function extractShareLinkFromBody() {
    return new Promise((resolve, reject) => {
        Office.context.mailbox.item.body.getAsync(
            Office.CoercionType.Html,
            (result) => {
                if (result.status !== Office.AsyncResultStatus.Succeeded) {
                    reject(new Error(result.error?.message || "Body read failed"));
                    return;
                }

                const html = result.value || "";

                // 1. Extract all href values
                const hrefRe = /href=["']([^"']+)["']/gi;
                let m;
                while ((m = hrefRe.exec(html)) !== null) {
                    const parsed = parseShareUrl(m[1]);
                    if (parsed.conversationId) { resolve(parsed); return; }
                }

                // 2. Fallback: scan all bare URLs in text/HTML
                const urlRe = /https?:\/\/[^\s"'<>)]+/gi;
                while ((m = urlRe.exec(html)) !== null) {
                    const parsed = parseShareUrl(m[0]);
                    if (parsed.conversationId) { resolve(parsed); return; }
                }

                resolve({ conversationId: null, docId: null });
            }
        );
    });
}

// ── Extract conversation_id and doc_id from a share URL ───────────────────
// Supports query-param formats:
//   ?conversation_id=xxx&doc_id=yyy
//   ?conversationId=xxx&documentId=yyy
//   ?cid=xxx&did=yyy
// And path formats:
//   /view/{conversationId}/{docId}
//   /share/{conversationId}
function parseShareUrl(rawUrl) {
    try {
        // Strip trailing punctuation that may have been captured
        const url = rawUrl.replace(/[>)"'\s]+$/, "");
        const u   = new URL(url);
        const sp  = u.searchParams;

        const conversationId =
            sp.get("conversation_id") ||
            sp.get("conversationId")  ||
            sp.get("cid")             || null;

        const docId =
            sp.get("doc_id")      ||
            sp.get("documentId")  ||
            sp.get("did")         || null;

        if (conversationId) return { conversationId, docId };

        // Path-based: /…/{conversationId}/{docId} or /…/{conversationId}
        const segments = u.pathname.split("/").filter(Boolean);
        if (segments.length >= 2) {
            return {
                conversationId: segments[segments.length - 2],
                docId:          segments[segments.length - 1]
            };
        }
        if (segments.length === 1) {
            return { conversationId: segments[0], docId: null };
        }
    } catch (_) { /* not a valid URL */ }

    return { conversationId: null, docId: null };
}

function setReadInitStatus(msg) {
    document.getElementById("read-init-status").textContent = msg;
    document.querySelector(".read-spinner-wrap").style.display = "flex";
    document.getElementById("read-init-error").classList.add("hidden");
    document.getElementById("read-init-status").classList.remove("hidden");
}

function showReadInitError(msg) {
    document.querySelector(".read-spinner-wrap").style.display = "none";
    document.getElementById("read-init-status").classList.add("hidden");
    document.getElementById("read-error-msg").textContent = msg;
    document.getElementById("read-init-error").classList.remove("hidden");
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
        // Step 1 — authenticate
        showComposeStatus("Signing in…");
        const token = await getAuthToken();

        // Step 2 — upload document
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

        const uploadData   = await uploadResp.json();
        // Prefer conversation_id; fall back to doc_id for the share payload
        const conversationId = uploadData.conversation_id || uploadData.conversationId || "";
        const docId          = uploadData.doc_id || uploadData.documentId || uploadData.id || "";

        if (!conversationId && !docId)
            throw new Error("No conversation ID or document ID returned by upload.");

        // Step 3 — share API (passes conversation_id + recipient list)
        showComposeStatus("Creating share link…");
        const shareLink = await callShareApi(token, conversationId, docId);

        // Step 4 — insert link into email body
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
// POST /v1/document/share
// Body: { conversation_id, doc_id (optional), sender_email, recipients: [...] }
// Returns: { share_url }
async function callShareApi(token, conversationId, docId) {
    const payload = {
        sender_email: _senderEmail,
        recipients:   _composeRecipients,
    };
    if (conversationId) payload.conversation_id = conversationId;
    if (docId)          payload.doc_id           = docId;

    const resp = await fetch(SHARE_URL, {
        method:  "POST",
        headers: {
            "Content-Type": "application/json",
            Authorization:  "Bearer " + token,
        },
        body: JSON.stringify(payload),
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
// ATTACHMENT BLOB HELPER  (compose upload)
// ══════════════════════════════════════════════════════════════════════════
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
// CHAT (shared by both modes via enterChat)
// ══════════════════════════════════════════════════════════════════════════
async function enterChat(conversationId, documentId, token) {
    currentConversationId = conversationId;
    currentDocumentId     = documentId;

    // Hide whichever init view is showing, reveal chat
    document.getElementById("view-read-init").classList.add("hidden");
    document.getElementById("view-compose").classList.add("hidden");
    document.getElementById("view-chat").classList.remove("hidden");
    document.getElementById("chat-history").innerHTML = "";
    hideSuggestions();

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
            appendMessage("ai", "Document ready. How can I help you?");
        }
    } catch (err) {
        hideTypingIndicator();
        appendMessage("ai", "Document ready. How can I help you?");
    }
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

// ══════════════════════════════════════════════════════════════════════════
// UTILITIES
// ══════════════════════════════════════════════════════════════════════════
function formatBytes(bytes) {
    if (!bytes) return "";
    if (bytes < 1024)    return bytes + " B";
    if (bytes < 1048576) return (bytes / 1024).toFixed(1) + " KB";
    return (bytes / 1048576).toFixed(1) + " MB";
}
