// ── Config ─────────────────────────────────────────────────────────────────
const PROXY_BASE       = "https://headphone-crust-stipulate.ngrok-free.dev";
const BLUE_BASE        = "https://demo.smartblue.ai";

// Domains whose /conversation URLs are treated as SmartBlue share links.
// Only emails containing a link on one of these domains will show "Start Chat".
const SMARTBLUE_DOMAINS = ["demo.smartblue.ai", "try.smartblue.ai", "app.smartblue.ai"];

const AUTH_URL         = `${PROXY_BASE}/v1/authenticate`;
const UPLOAD_URL       = `${PROXY_BASE}/v1/document/upload`;
const BUNDLE_ADD_URL   = `${PROXY_BASE}/v1/document/bundle/add`;
const SHARE_URL        = `${PROXY_BASE}/v1/document/share`;
const WELCOME_URL      = `${PROXY_BASE}/v1/conversation/ask/welcome`;
const ASK_URL          = `${PROXY_BASE}/v1/conversation/ask/question`;
const CONVERSATION_URL = `${PROXY_BASE}/v1/conversation`;

const AZURE_CLIENT_ID  = "c49037f2-0565-4a5c-8b17-f9b8b3ee35c7";
const AZURE_TENANT_ID  = "f895e126-dbc8-41bb-b00b-5cd2172346f9";
const SCOPES           = ["openid", "profile", "email", "User.Read"];

const msalConfig = {
    auth: {
        clientId:    AZURE_CLIENT_ID,
        authority:   "https://login.microsoftonline.com/" + AZURE_TENANT_ID,
        redirectUri: window.location.href.split("?")[0],
    },
    cache: { cacheLocation: "sessionStorage", storeAuthStateInCookie: false },
};

// ── State ──────────────────────────────────────────────────────────────────
const state = { currentConversationId: null, currentDocumentId: null };

let _msal                 = null;
let _cachedSmartBlueToken = null;
let _customProps          = null;
let _composeAttachments   = [];
let _composeRecipients    = [];
let _senderEmail          = "";
let _readShareInfo        = null;

// ── MSAL / Auth ────────────────────────────────────────────────────────────
function getMsal() {
    if (!_msal) _msal = new msal.PublicClientApplication(msalConfig);
    return _msal;
}
function clearToken() { _cachedSmartBlueToken = null; }

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
        } catch (e) { throw new Error("Sign-in failed: " + (e.message || e.errorCode)); }
    }
    const authResp = await fetch(AUTH_URL, {
        method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ idToken }),
    });
    if (!authResp.ok) throw new Error("Auth failed (" + authResp.status + "): " + await authResp.text());
    const { token } = await authResp.json();
    if (!token) throw new Error("No token returned from auth proxy.");
    _cachedSmartBlueToken = token;
    return token;
}

// ── Custom Properties ──────────────────────────────────────────────────────
const PROP_KEY = "conversationsMap";

function loadCustomProps() {
    return new Promise((resolve, reject) => {
        Office.context.mailbox.item.loadCustomPropertiesAsync(result => {
            if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
            else reject(new Error(result.error?.message || "Failed to load custom properties"));
        });
    });
}
function getConversationMap(cp) {
    try { const raw = cp.get(PROP_KEY); return raw ? JSON.parse(raw) : {}; } catch { return {}; }
}
function saveConversationRecord(cp, fingerprint, record) {
    const map = getConversationMap(cp);
    map[fingerprint] = record;
    cp.set(PROP_KEY, JSON.stringify(map));
    return new Promise((resolve, reject) => {
        cp.saveAsync(result => {
            if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
            else reject(new Error(result.error?.message || "Failed to save custom properties"));
        });
    });
}
function singleFingerprint(att) { return `att_${att.id}`; }
function bundleFingerprint(primaryAtt, secondaryAtts) {
    const ids = secondaryAtts.map(a => a.id).sort();
    return `bundle_${[primaryAtt.id, ...ids].join("_")}`;
}

// ── URL parsing ────────────────────────────────────────────────────────────
function parseDocUrl(rawUrl) {
    try {
        const url = rawUrl.replace(/[>)"'\s]+$/, "").replace(/&amp;/gi, "&");
        const u   = new URL(url);

        // Only treat URLs from known SmartBlue domains as share links
        if (!SMARTBLUE_DOMAINS.includes(u.hostname)) {
            return { conversationId: null, docId: null, shareUrl: null };
        }

        const sp = u.searchParams;
        const conversationId =
            sp.get("conversation-id") || sp.get("conversation_id") ||
            sp.get("conversationId")  || sp.get("cid") || null;
        const docId =
            sp.get("doc-id") || sp.get("doc_id") || sp.get("documentId") || sp.get("did") || null;

        if (conversationId) return { conversationId, docId, shareUrl: url };
    } catch (_) {}
    return { conversationId: null, docId: null, shareUrl: null };
}

function extractShareLinkFromBody() {
    return new Promise((resolve, reject) => {
        Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, (result) => {
            if (result.status !== Office.AsyncResultStatus.Succeeded) {
                reject(new Error(result.error?.message || "Body read failed")); return;
            }
            const html = result.value || "";
            const anchorRe = /<a[^>]+href=["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>/gi;
            let m;
            while ((m = anchorRe.exec(html)) !== null) {
                const linkText = m[2].replace(/<[^>]+>/g, "").replace(/\s+/g, " ").trim();
                const parsed   = parseDocUrl(m[1]);
                if (parsed.conversationId) { resolve({ ...parsed, linkText: linkText || m[1] }); return; }
            }
            const urlRe = /https?:\/\/[^\s"'<>)]+/gi;
            while ((m = urlRe.exec(html)) !== null) {
                const parsed = parseDocUrl(m[0]);
                if (parsed.conversationId) { resolve({ ...parsed, linkText: null }); return; }
            }
            resolve({ conversationId: null, docId: null, linkText: null });
        });
    });
}

// ── Utilities ──────────────────────────────────────────────────────────────
function escHtml(str) {
    return (str || "")
        .replace(/&/g, "&amp;").replace(/</g, "&lt;")
        .replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
function formatBytes(bytes) {
    if (!bytes) return "";
    if (bytes < 1024)    return bytes + " B";
    if (bytes < 1048576) return (bytes / 1024).toFixed(1) + " KB";
    return (bytes / 1048576).toFixed(1) + " MB";
}
function formatResponse(raw, conversationId, documentId) {
    const doc_url = `${BLUE_BASE}/conversation?conversation-id=${conversationId}&doc-id=${documentId}`;
    let text = raw.replace(
        /<blueEmbed-doc-page>[^:]+:[^:]+:(\d+)<\/blueEmbed-doc-page>/g,
        `<a href="${doc_url}" target="_blank" class="page-ref" data-page="$1">pg $1</a>`
    );
    text = text.replace(/\*\*(.+?)\*\*/g, "<strong>$1</strong>");
    const lines = text.split(/\n/);
    let html = "", inList = false;
    for (const rawLine of lines) {
        const line = rawLine.trim();
        if (!line) { if (inList) { html += "</ul>"; inList = false; } continue; }
        if (/^[*\u25CF\u2022]\s+/.test(line)) {
            if (!inList) { html += '<ul class="ai-list">'; inList = true; }
            html += "<li>" + line.replace(/^[*\u25CF\u2022]\s+/, "") + "</li>";
        } else {
            if (inList) { html += "</ul>"; inList = false; }
            html += "<p>" + line + "</p>";
        }
    }
    if (inList) html += "</ul>";
    return html;
}
function getMimeType(filename) {
    const ext = (filename || "").split(".").pop().toLowerCase();
    const MAP = {
        pdf:"application/pdf", doc:"application/msword",
        docx:"application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        xls:"application/vnd.ms-excel",
        xlsx:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        csv:"text/csv", ppt:"application/vnd.ms-powerpoint",
        pptx:"application/vnd.openxmlformats-officedocument.presentationml.presentation",
        txt:"text/plain", rtf:"application/rtf", png:"image/png",
        jpg:"image/jpeg", jpeg:"image/jpeg", gif:"image/gif", webp:"image/webp", zip:"application/zip",
    };
    return MAP[ext] || "application/octet-stream";
}
function fallbackCopy(text, cb) {
    const ta = document.createElement("textarea");
    ta.value = text; ta.style.cssText = "position:fixed;opacity:0";
    document.body.appendChild(ta); ta.select();
    try { document.execCommand("copy"); cb(); } catch (_) {}
    document.body.removeChild(ta);
}

// ── Upload ─────────────────────────────────────────────────────────────────
function getAttachmentBlob(attachmentId, filename) {
    return new Promise((resolve, reject) => {
        Office.context.mailbox.item.getAttachmentContentAsync(attachmentId, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                const binary = atob(result.value.content);
                const bytes  = new Uint8Array(binary.length);
                for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
                resolve(new Blob([bytes], { type: getMimeType(filename) }));
            } else { reject(new Error(result.error.message)); }
        });
    });
}
async function uploadPrimary(att, token) {
    const blob = await getAttachmentBlob(att.id, att.name);
    const form = new FormData();
    form.append("document", blob, att.name);
    const resp = await fetch(UPLOAD_URL, { method:"POST", headers:{Authorization:"Bearer "+token}, body:form });
    if (!resp.ok) throw new Error("Upload failed (" + resp.status + "): " + await resp.text());
    const data = await resp.json();
    console.log("Upload response:", JSON.stringify(data));
    const conversationId = data.conversation_id || data.conversationId || null;
    const documentId = data.doc_id || data.document_id || data.documentId || data.docId || data.doc || data.id || null;
    if (!conversationId) throw new Error("No conversation_id returned by upload.");
    if (!documentId)     throw new Error("No document ID returned by upload (keys: " + Object.keys(data).join(", ") + ").");
    return { conversationId, documentId };
}
async function uploadSupportingById(att, conversationId, token) {
    const blob = await getAttachmentBlob(att.id, att.name);
    const form = new FormData();
    form.append("document", blob, att.name);
    const resp = await fetch(`${BUNDLE_ADD_URL}?conversation_id=${encodeURIComponent(conversationId)}`,
        { method:"POST", headers:{Authorization:"Bearer "+token}, body:form });
    if (!resp.ok) console.warn("Supporting upload failed:", att.name, await resp.text());
}

// ── Remove attachment helper ──────────────────────────────────────────────
// Only works for attachments the add-in can control; user-added attachments
// will be rejected by Outlook — the error is caught and warned, not thrown.
async function removeAttachmentIfRequested(attachmentIds) {
    const checkbox = document.getElementById("chk-include-attachment");
    if (!checkbox || !checkbox.checked) return;
    const ids = Array.isArray(attachmentIds) ? attachmentIds : [attachmentIds];
    for (const id of ids) {
        await new Promise((resolve) => {
            Office.context.mailbox.item.removeAttachmentAsync(id, (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    console.log("Attachment removed:", id);
                } else {
                    console.warn("Remove attachment failed (non-fatal):", id, result.error?.message);
                }
                resolve();
            });
        });
    }
}

// ── API calls ──────────────────────────────────────────────────────────────
async function callShareApi(token, conversationId, docId, senderEmail, recipients) {
    const payload = { sender_email: senderEmail, recipients };
    if (conversationId) payload.conversation_id = conversationId;
    if (docId)          payload.doc_id           = docId;
    const resp = await fetch(SHARE_URL, {
        method:"POST", headers:{"Content-Type":"application/json", Authorization:"Bearer "+token},
        body: JSON.stringify(payload),
    });
    if (!resp.ok) throw new Error("Share API failed (" + resp.status + "): " + await resp.text());
    const data = await resp.json();
    const url  = data.share_url || data.shareUrl || data.url || "";
    if (!url) throw new Error("Share API returned no URL.");
    return url;
}
function fetchHistory(token, conversationId) {
    // Matches server route: GET /v1/conversation/history?conversation_id={id}
    return fetch(`${CONVERSATION_URL}/history?conversation_id=${encodeURIComponent(conversationId)}`, {
        headers: { Authorization: "Bearer " + token, "ngrok-skip-browser-warning": "true" },
    });
}
function fetchWelcome(token, conversationId, documentId) {
    return fetch(WELCOME_URL, {
        method:"POST",
        headers: { "Content-Type":"application/json", Authorization:"Bearer "+token, "ngrok-skip-browser-warning":"true" },
        body: JSON.stringify({ conversationId, documentId }),
    });
}
function askQuestion(token, conversationId, text) {
    return fetch(ASK_URL, {
        method:"POST",
        headers: { "Content-Type":"application/json", Authorization:"Bearer "+token },
        body: JSON.stringify({ conversationId, text, isMobile:false }),
    });
}

// ── Status ─────────────────────────────────────────────────────────────────
function showReadStatus(msg)    { const el = document.getElementById("status-msg");    if (el) el.innerText = msg; }
function showComposeStatus(msg) { document.getElementById("compose-status").innerText = msg; }
function showReadInitError(msg) {
    document.querySelector(".read-spinner-wrap").style.display = "none";
    document.getElementById("read-init-status").classList.add("hidden");
    document.getElementById("read-error-msg").textContent = msg;
    document.getElementById("read-init-error").classList.remove("hidden");
}

// ── Chat UI ────────────────────────────────────────────────────────────────
function showTypingIndicator() {
    const hist = document.getElementById("chat-history");
    if (hist.querySelector(".msg-typing")) return;
    const div = document.createElement("div");
    div.className = "msg-typing"; div.id = "typing-indicator";
    div.innerHTML = `<span class="typing-dot"></span><span class="typing-dot"></span><span class="typing-dot"></span>`;
    hist.appendChild(div); hist.scrollTop = hist.scrollHeight;
}
function hideTypingIndicator() { const el = document.getElementById("typing-indicator"); if (el) el.remove(); }
function appendMessage(role, text, conversationId, documentId) {
    const hist = document.getElementById("chat-history");
    const div  = document.createElement("div");
    if (role === "user") {
        div.className = "msg-user";
        const p = document.createElement("p"); p.textContent = text; div.appendChild(p);
    } else { div.className = "msg-ai"; div.innerHTML = formatResponse(text, conversationId, documentId); }
    hist.appendChild(div); hist.scrollTop = hist.scrollHeight;
}
function hideSuggestions() { const box = document.getElementById("suggestions"); box.classList.add("hidden"); box.innerHTML = ""; }
function renderSuggestions(tags) {
    const box = document.getElementById("suggestions");
    box.innerHTML = "";
    tags.forEach(tag => {
        const q = typeof tag === "string" ? tag : (tag["next-question"] || tag.question || "");
        if (!q.trim()) return;
        const chip = document.createElement("button");
        chip.className = "chip"; chip.textContent = q;
        chip.onclick = () => { hideSuggestions(); document.getElementById("user-input").value = q; sendChatMessage(); };
        box.appendChild(chip);
    });
    box.classList.remove("hidden");
}
function restoreConversationHistory(messages, conversationId, documentId) {
    messages.forEach(msg => {
        const text = msg.text || "";
        if (!text.trim()) return;
        appendMessage(msg.sender === "assistant" ? "ai" : "user", text, conversationId, documentId);
    });
    const lastAI = [...messages].reverse().find(m => m.sender === "assistant");
    if (lastAI && Array.isArray(lastAI.tags) && lastAI.tags.length) renderSuggestions(lastAI.tags);
    document.getElementById("chat-history").scrollTop = document.getElementById("chat-history").scrollHeight;
}

// ── Attachment UI ──────────────────────────────────────────────────────────
function renderBundleList(attachments, container) {
    container.innerHTML = "";
    attachments.forEach((att, index) => {
        const isPrimary = index === 0;
        const div = document.createElement("div");
        div.className = "att-item" + (isPrimary ? " is-primary" : "");
        div.dataset.index = index;
        div.innerHTML = `
            <div class="att-bundle-row">
                <div class="att-radio-col">
                    <input type="radio" name="primaryIndex" value="${index}" id="radio-${index}" ${isPrimary ? "checked" : ""}/>
                    <label class="radio-label" for="radio-${index}">Primary</label>
                </div>
                <div class="att-info">
                    <div class="att-name" title="${escHtml(att.name)}">${escHtml(att.name)}</div>
                    <div class="att-meta">${formatBytes(att.size)}</div>
                </div>
                <div class="att-secondary-col">
                    <input type="checkbox" name="secondaryIndex" value="${index}" id="chk-sec-${index}"
                           ${isPrimary ? "" : "checked"} ${isPrimary ? "disabled" : ""}/>
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
        const idx = item.dataset.index; const isPrimary = idx === primaryVal;
        const secChk = item.querySelector("input[name='secondaryIndex']");
        item.classList.toggle("is-primary", isPrimary);
        if (isPrimary) { secChk.checked = false; secChk.disabled = true; }
        else { secChk.disabled = false; if (!secChk.dataset.userUnchecked) secChk.checked = true; }
    });
}
document.addEventListener("change", (e) => {
    if (e.target.name === "secondaryIndex") e.target.dataset.userUnchecked = e.target.checked ? "" : "1";
});
function renderIndividualReadList(attachments, container, hasContext = false) {
    container.innerHTML = "";
    attachments.forEach((att, index) => {
        const div = document.createElement("div");
        div.className = "att-item";
        const btnLabel = hasContext ? "Add to Bundle" : "Upload";
        div.innerHTML = `
            <div class="att-individual-row">
                <div class="att-info">
                    <div class="att-name" title="${escHtml(att.name)}">${escHtml(att.name)}</div>
                    <div class="att-meta">${formatBytes(att.size)}</div>
                </div>
                <button class="btn-upload-single" data-index="${index}">${btnLabel}</button>
            </div>`;
        container.appendChild(div);
    });
    container.querySelectorAll(".btn-upload-single").forEach(btn => {
        btn.onclick = hasContext
            ? () => handleReadAddToExisting(parseInt(btn.dataset.index))
            : () => handleReadSingleUpload(parseInt(btn.dataset.index));
    });
}
function renderIndividualComposeList(attachments, container) {
    attachments.forEach((att, index) => {
        const div = document.createElement("div");
        div.className = "att-item";
        div.innerHTML = `
            <div class="att-individual-row">
                <div class="att-info">
                    <div class="att-name" title="${escHtml(att.name)}">${escHtml(att.name)}</div>
                    <div class="att-meta">${formatBytes(att.size)}</div>
                </div>
                <button class="btn-upload-single btn-upload-share" data-index="${index}">Share</button>
            </div>`;
        container.appendChild(div);
    });
    container.querySelectorAll(".btn-upload-share").forEach(btn => {
        btn.onclick = () => handleComposeSingleUpload(parseInt(btn.dataset.index));
    });
}

// ── Share section ──────────────────────────────────────────────────────────
function renderShareSection(shareInfo) {
    const section = document.getElementById("read-share-section");
    const card    = document.getElementById("read-share-card");
    const displayText = shareInfo.linkText || shareInfo.shareUrl || "View on SmartBlue";
    const displayUrl  = shareInfo.shareUrl  || "";
    card.innerHTML = `
        <div class="read-share-inner">
            <svg class="read-share-file-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                <polyline points="14 2 14 8 20 8"/>
            </svg>
            <div class="read-share-info">
                <div class="read-share-name" title="${escHtml(displayText)}">${escHtml(displayText)}</div>
                <div class="read-share-url" title="${escHtml(displayUrl)}">${escHtml(displayUrl)}</div>
            </div>
        </div>
        <button class="btn-start-chat" id="btn-share-chat">
            <svg width="12" height="12" viewBox="0 0 24 24" fill="currentColor"><polygon points="5 3 19 12 5 21 5 3"/></svg>
            Start Chat
        </button>`;
    section.classList.remove("hidden");
    document.getElementById("btn-share-chat").onclick = async () => {
        const btn = document.getElementById("btn-share-chat");
        btn.disabled = true;
        showReadStatus("Signing in\u2026");
        try {
            const token = await getAuthToken();
            await enterChat(shareInfo.conversationId, shareInfo.docId, token);
        } catch (err) { showReadStatus("Error: " + err.message); clearToken(); btn.disabled = false; }
    };
}
function insertShareLinkIntoBody(link, filename) {
    return new Promise((resolve) => {
        const html = `<p style="font-family:sans-serif;margin:8px 0;">`
                   + `<a href="${link}" target="_blank" style="color:#0D47A1;font-size:14px;font-weight:500;text-decoration:none;">`
                   + `${filename} \u2014 View on SmartBlue</a></p>`
                   + `<p style="font-family:sans-serif;margin:8px 0;">Access the document directly in BlueAI using the link below for a secure and seamless viewing experience.</p>`;
        Office.context.mailbox.item.body.setSelectedDataAsync(html, { coercionType:Office.CoercionType.Html }, (result) => {
            if (result.status === Office.AsyncResultStatus.Failed) {
                Office.context.mailbox.item.body.setSelectedDataAsync(`\n${filename}: ${link}\n`,
                    { coercionType:Office.CoercionType.Text }, () => resolve());
            } else { resolve(); }
        });
    });
}
function renderComposeResult(link) {
    document.getElementById("result-link-text").textContent = link;
    document.getElementById("compose-result").classList.remove("hidden");
    document.getElementById("compose-result").scrollIntoView({ behavior:"smooth" });
}

// ══════════════════════════════════════════════════════════════════════════
// ENTRY POINT
// ══════════════════════════════════════════════════════════════════════════
Office.onReady((info) => { if (info.host === Office.HostType.Outlook) init(); });

function init() {
    const item = Office.context.mailbox.item;
    const isCompose = typeof item.subject?.setAsync === "function"
                   || typeof item.body?.setAsync    === "function";
    if (isCompose) { initCompose(); } else { initRead(); }
}

// ══════════════════════════════════════════════════════════════════════════
// READ MODE
// ══════════════════════════════════════════════════════════════════════════
function initRead() {
    document.querySelector(".header-title").textContent = "View Document";
    document.getElementById("btn-send").onclick          = sendChatMessage;
    document.getElementById("user-input").addEventListener("keydown", (e) => {
        if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); sendChatMessage(); }
    });
    document.getElementById("btn-back").onclick          = switchToReadView;
    document.getElementById("btn-upload-bundle").onclick = handleReadBundleUpload;
    document.getElementById("chk-bulk").onchange         = onReadToggleMode;

    document.getElementById("view-read-init").classList.remove("hidden");

    extractShareLinkFromBody()
        .then(shareInfo => {
            document.getElementById("view-read-init").classList.add("hidden");
            document.getElementById("view-read").classList.remove("hidden");
            if (shareInfo && shareInfo.conversationId) {
                _readShareInfo = shareInfo;
                renderShareSection(shareInfo);
            }

            loadCustomProps()
                .then(cp => {
                    _customProps = cp;

                    // Auto-open chat if we have enough context — no manual step needed:
                    // Priority 1: shared URL in email body (has conversationId + docId)
                    // Priority 2: most recent previous chat from custom props
                    if (shareInfo && shareInfo.conversationId && shareInfo.docId) {
                        getAuthToken()
                            .then(token => enterChat(shareInfo.conversationId, shareInfo.docId, token))
                            .catch(err => {
                                console.warn("Auto-open from share link failed:", err.message);
                                renderPreviousChats();
                            });
                        return; // skip rendering prev chats / attachments — enterChat shows the chat view
                    }

                    const records = Object.values(getConversationMap(cp))
                        .sort((a, b) => (b.timestamp || 0) - (a.timestamp || 0));

                    if (records.length > 0) {
                        // Auto-resume the most recent conversation
                        const latest = records[0];
                        getAuthToken()
                            .then(token => enterChat(latest.conversationId, latest.documentId, token))
                            .catch(err => {
                                console.warn("Auto-resume failed:", err.message);
                                renderPreviousChats();
                            });
                        return;
                    }

                    // No context — show normal read view
                    renderPreviousChats();
                })
                .catch(() => {
                    // Custom props unavailable — fall through to normal UI
                    if (shareInfo && shareInfo.conversationId && shareInfo.docId) {
                        getAuthToken()
                            .then(token => enterChat(shareInfo.conversationId, shareInfo.docId, token))
                            .catch(err => console.warn("Auto-open failed:", err.message));
                    }
                });

            loadReadAttachments();
            const atts = Office.context.mailbox.item.attachments || [];
            if (shareInfo && shareInfo.conversationId && atts.length > 0)
                document.getElementById("read-or-divider").classList.remove("hidden");
        })
        .catch(err => showReadInitError("Error reading email: " + err.message));
}
function isReadBulkMode() { return document.getElementById("chk-bulk").checked; }
function onReadToggleMode() {
    const bulk = isReadBulkMode();
    document.getElementById("lbl-bundle").classList.toggle("active", bulk);
    document.getElementById("lbl-individual").classList.toggle("active", !bulk);
    document.getElementById("bundle-footer").classList.toggle("hidden", !bulk);
    loadReadAttachments();
}
function renderPreviousChats() {
    const section = document.getElementById("read-prev-section");
    const list    = document.getElementById("read-prev-list");
    if (!_customProps) { section.classList.add("hidden"); return; }
    const records = Object.values(getConversationMap(_customProps));
    if (!records.length) { section.classList.add("hidden"); return; }
    list.innerHTML = "";
    records.slice().sort((a, b) => (b.timestamp || 0) - (a.timestamp || 0)).forEach(rec => {
        const item = document.createElement("div");
        item.className = "prev-chat-item";
        const date = rec.timestamp
            ? new Date(rec.timestamp).toLocaleDateString(undefined, { month:"short", day:"numeric" }) : "";
        item.innerHTML = `
            <div class="prev-chat-info">
                <div class="prev-chat-name" title="${escHtml(rec.label || "")}">${escHtml(rec.label || "Document")}</div>
                <div class="prev-chat-meta">${escHtml(rec.uploadType || "")}${date ? " \u00b7 " + date : ""}</div>
            </div>
            <button class="btn-resume">Resume</button>`;
        item.querySelector(".btn-resume").onclick = async () => {
            showReadStatus("Signing in\u2026");
            try { const token = await getAuthToken(); await enterChat(rec.conversationId, rec.documentId, token); }
            catch (err) { showReadStatus("Error: " + err.message); clearToken(); }
        };
        list.appendChild(item);
    });
    section.classList.remove("hidden");
}
function loadReadAttachments() {
    const attachments   = Office.context.mailbox.item.attachments || [];
    const listDiv       = document.getElementById("attachment-list");
    const footerDiv     = document.getElementById("bundle-footer");
    const attachSection = document.getElementById("read-attach-section");
    const divider       = document.getElementById("read-or-divider");
    if (!attachments.length) {
        // No attachments — hide the entire section and the divider above it
        if (attachSection) attachSection.classList.add("hidden");
        if (divider)       divider.classList.add("hidden");
        return;
    }
    // Has attachments — make sure section is visible
    if (attachSection) attachSection.classList.remove("hidden");
    const hasContext = !!((_readShareInfo && _readShareInfo.conversationId) ||
        (_customProps && Object.keys(getConversationMap(_customProps)).length > 0));

    if (isReadBulkMode()) {
        footerDiv.classList.remove("hidden");
        const bundleBtn = document.getElementById("btn-upload-bundle");
        bundleBtn.disabled = false;
        if (hasContext) {
            bundleBtn.textContent = "＋ Add to Bundle";
            bundleBtn.onclick = handleReadAddToBundle;
        } else {
            bundleBtn.textContent = "⬆ Upload & Analyse";
            bundleBtn.onclick = handleReadBundleUpload;
        }
        renderBundleList(attachments, listDiv);
    } else {
        footerDiv.classList.add("hidden");
        renderIndividualReadList(attachments, listDiv, hasContext);
    }
}
async function handleReadBundleUpload() {
    const attachments  = Office.context.mailbox.item.attachments;
    const primaryRadio = document.querySelector("input[name='primaryIndex']:checked");
    if (!primaryRadio) { showReadStatus("Please select a primary document."); return; }
    const primaryIndex  = parseInt(primaryRadio.value);
    const primaryAtt    = attachments[primaryIndex];
    const secondaryAtts = Array.from(document.querySelectorAll("input[name='secondaryIndex']:checked"))
        .map(c => parseInt(c.value)).filter(i => i !== primaryIndex).map(i => attachments[i]);
    document.getElementById("btn-upload-bundle").disabled = true;
    showReadStatus("Signing in\u2026");
    try {
        const token = await getAuthToken();
        if (_customProps) {
            const fp = bundleFingerprint(primaryAtt, secondaryAtts);
            const rec = getConversationMap(_customProps)[fp];
            if (rec) { showReadStatus(""); return await enterChat(rec.conversationId, rec.documentId, token); }
        }
        showReadStatus("Uploading primary document\u2026");
        const { conversationId, documentId } = await uploadPrimary(primaryAtt, token);
        if (secondaryAtts.length > 0) {
            showReadStatus("Uploading " + secondaryAtts.length + " supporting doc(s)\u2026");
            for (const att of secondaryAtts) await uploadSupportingById(att, conversationId, token);
        }
        if (_customProps) {
            saveConversationRecord(_customProps, bundleFingerprint(primaryAtt, secondaryAtts), {
                conversationId, documentId, label: primaryAtt.name, uploadType:"bundle", timestamp:Date.now(),
            }).catch(err => console.warn("customProps save failed:", err.message));
        }
        await enterChat(conversationId, documentId, token);
    } catch (err) {
        console.error("Read bundle upload error:", err); showReadStatus("Error: " + err.message); clearToken();
        document.getElementById("btn-upload-bundle").disabled = false;
    }
}
async function handleReadSingleUpload(index) {
    const att = Office.context.mailbox.item.attachments[index];
    document.querySelectorAll(".btn-upload-single").forEach(b => b.disabled = true);
    showReadStatus("Signing in\u2026");
    try {
        const token = await getAuthToken();
        if (_customProps) {
            const fp = singleFingerprint(att);
            const rec = getConversationMap(_customProps)[fp];
            if (rec) { showReadStatus(""); return await enterChat(rec.conversationId, rec.documentId, token); }
        }
        showReadStatus("Uploading " + att.name + "\u2026");
        const { conversationId, documentId } = await uploadPrimary(att, token);
        if (_customProps) {
            saveConversationRecord(_customProps, singleFingerprint(att), {
                conversationId, documentId, label:att.name, uploadType:"single", timestamp:Date.now(),
            }).catch(err => console.warn("customProps save failed:", err.message));
        }
        await enterChat(conversationId, documentId, token);
    } catch (err) {
        console.error("Read single upload error:", err); showReadStatus("Error: " + err.message); clearToken();
        document.querySelectorAll(".btn-upload-single").forEach(b => b.disabled = false);
    }
}
// Upload selected attachments as supporting docs to the existing conversation.
// Called from bundle footer when existing context is present.
async function handleReadAddToBundle() {
    const attachments   = Office.context.mailbox.item.attachments;
    const secondaryAtts = Array.from(document.querySelectorAll("input[name='secondaryIndex']:checked"))
        .map(c => parseInt(c.value)).map(i => attachments[i]);
    if (!secondaryAtts.length) { showReadStatus("Select at least one document to add."); return; }

    const existingConvId = _readShareInfo?.conversationId
        || Object.values(getConversationMap(_customProps || {})).sort((a,b)=>(b.timestamp||0)-(a.timestamp||0))[0]?.conversationId;
    if (!existingConvId) { showReadStatus("No existing conversation found."); return; }

    document.getElementById("btn-upload-bundle").disabled = true;
    showReadStatus("Signing in\u2026");
    try {
        const token = await getAuthToken();
        showReadStatus("Adding " + secondaryAtts.length + " doc(s) to conversation\u2026");
        for (const att of secondaryAtts) {
            await uploadSupportingById(att, existingConvId, token);
        }
        showReadStatus("");
        await enterChat(existingConvId, _readShareInfo?.docId
            || Object.values(getConversationMap(_customProps||{})).sort((a,b)=>(b.timestamp||0)-(a.timestamp||0))[0]?.documentId,
            token);
    } catch (err) {
        console.error("Add to bundle error:", err); showReadStatus("Error: " + err.message); clearToken();
        document.getElementById("btn-upload-bundle").disabled = false;
    }
}

// Add a single attachment as a supporting doc to the existing conversation.
// Called from individual list buttons when existing context is present.
async function handleReadAddToExisting(index) {
    const att = Office.context.mailbox.item.attachments[index];
    document.querySelectorAll(".btn-upload-single").forEach(b => b.disabled = true);
    showReadStatus("Signing in\u2026");
    try {
        const token = await getAuthToken();
        const existingConvId = _readShareInfo?.conversationId
            || Object.values(getConversationMap(_customProps || {})).sort((a,b)=>(b.timestamp||0)-(a.timestamp||0))[0]?.conversationId;
        if (!existingConvId) throw new Error("No existing conversation found.");
        showReadStatus("Adding " + att.name + " to conversation\u2026");
        await uploadSupportingById(att, existingConvId, token);
        showReadStatus("");
        const existingDocId = _readShareInfo?.docId
            || Object.values(getConversationMap(_customProps||{})).sort((a,b)=>(b.timestamp||0)-(a.timestamp||0))[0]?.documentId;
        await enterChat(existingConvId, existingDocId, token);
    } catch (err) {
        console.error("Add to existing error:", err); showReadStatus("Error: " + err.message); clearToken();
        document.querySelectorAll(".btn-upload-single").forEach(b => b.disabled = false);
    }
}

function switchToReadView() {
    document.getElementById("view-chat").classList.add("hidden");
    document.getElementById("view-read").classList.remove("hidden");
    document.getElementById("btn-back").classList.add("hidden");
    document.getElementById("chat-history").innerHTML = "";
    hideSuggestions();
    state.currentConversationId = null; state.currentDocumentId = null;
    _readShareInfo = null;  // reset so back-navigation works correctly
    const shareBtn = document.getElementById("btn-share-chat");
    if (shareBtn) shareBtn.disabled = false;
    renderPreviousChats(); 
    loadReadAttachments(); 
    showReadStatus("");
}

// ══════════════════════════════════════════════════════════════════════════
// COMPOSE MODE
// ══════════════════════════════════════════════════════════════════════════
function initCompose() {
    document.querySelector(".header-title").textContent = "Share Document";
    document.getElementById("view-compose").classList.remove("hidden");
    document.getElementById("btn-refresh").classList.remove("hidden");
    document.getElementById("btn-refresh").onclick        = () => { state.suppressAttachmentRefresh = false; loadComposeData(true); };
    document.getElementById("btn-compose-upload").onclick = handleComposeBundleUpload;
    document.getElementById("btn-copy-link").onclick      = copyResultLink;
    document.getElementById("chk-compose-bulk").onchange  = onComposeToggleMode;
    loadComposeData(false);

    // Live sync — fire loadComposeData whenever the user adds/removes
    // an attachment or changes recipients, removing the need to manually refresh.
    // Requires Mailbox 1.8 (AttachmentsChanged) / 1.7 (RecipientsChanged).
    // The refresh button stays as a fallback for older clients.
    if (Office.context.requirements.isSetSupported("Mailbox", "1.8")) {
        Office.context.mailbox.item.addHandlerAsync(
            Office.EventType.AttachmentsChanged,
            () => loadComposeData(true)
        );
    }
    if (Office.context.requirements.isSetSupported("Mailbox", "1.7")) {
        Office.context.mailbox.item.addHandlerAsync(
            Office.EventType.RecipientsChanged,
            () => loadComposeData(true)
        );
    }
}
function isComposeBulkMode() { return document.getElementById("chk-compose-bulk").checked; }
function onComposeToggleMode() {
    const bulk = isComposeBulkMode();
    // Elements may be absent if the HTML hides/removes them — use optional chaining
    document.getElementById("clbl-bundle")?.classList.toggle("active", bulk);
    document.getElementById("clbl-individual")?.classList.toggle("active", !bulk);
    document.getElementById("compose-bundle-footer").classList.toggle("hidden", !bulk);
    renderComposeAttachments(_composeAttachments);
}
function loadComposeData(isRefresh) {
    if (isRefresh) {
        document.getElementById("btn-refresh").classList.add("spinning");
        document.getElementById("compose-result").classList.add("hidden");
        showComposeStatus("");
    }
    try { _senderEmail = Office.context.mailbox.userProfile.emailAddress || ""; } catch (e) { _senderEmail = ""; }
    const item = Office.context.mailbox.item;
    Promise.all([
        new Promise(res => item.to.getAsync(r => res(r.status === Office.AsyncResultStatus.Succeeded ? r.value : []))),
        new Promise(res => item.cc.getAsync(r => res(r.status === Office.AsyncResultStatus.Succeeded ? r.value : []))),
        new Promise(res => item.bcc.getAsync(r => res(r.status === Office.AsyncResultStatus.Succeeded ? r.value : []))),
        new Promise(res => item.getAttachmentsAsync(r => res(r.status === Office.AsyncResultStatus.Succeeded ? r.value : [])))
    ]).then(([toList, ccList, bccList, attachments]) => {
        const seen = new Set();
        _composeRecipients = [...toList, ...ccList, ...bccList]
            .map(r => (r.emailAddress || "").toLowerCase().trim())
            .filter(e => { if (!e || seen.has(e)) return false; seen.add(e); return true; });
        renderComposeRecipients(toList, ccList, bccList);
        _composeAttachments = attachments;

        // Auto-select mode based on attachment count:
        // 1 attachment → Individual, 2+ → Bundle
        const bulkChk = document.getElementById("chk-compose-bulk");
        if (attachments.length === 1) {
            bulkChk.checked = false;
        } else if (attachments.length > 1) {
            bulkChk.checked = true;
        }
        onComposeToggleMode();  // sync labels + footer visibility

        renderComposeAttachments(_composeAttachments);
        document.getElementById("btn-refresh").classList.remove("spinning");
    });
}
function renderComposeRecipients(toList, ccList, bccList = []) {
    const area  = document.getElementById("compose-recipients");
    const badge = document.getElementById("recipients-count");
    const total = toList.length + ccList.length + bccList.length;
    badge.textContent = total || "";
    if (total === 0) { area.innerHTML = `<div class="compose-empty">No recipients yet. Add To / CC addresses then click &#8635; Refresh.</div>`; return; }
    area.innerHTML = "";
    const buildRow = (label, list) => {
        if (!list.length) return;
        const row = document.createElement("div"); row.className = "recipient-row";
        const lbl = document.createElement("span"); lbl.className = "recipient-row-label"; lbl.textContent = label;
        row.appendChild(lbl);
        const chips = document.createElement("div"); chips.className = "recipient-chips";
        list.forEach(r => {
            const chip = document.createElement("span"); chip.className = "recipient-chip";
            chip.textContent = r.emailAddress || r.displayName || ""; chip.title = r.emailAddress || "";
            chips.appendChild(chip);
        });
        row.appendChild(chips); area.appendChild(row);
    };
    buildRow("To:", toList); buildRow("CC:", ccList); buildRow("BCC:", bccList);
}
function renderComposeAttachments(attachments) {
    const list  = document.getElementById("compose-attachments");
    const badge = document.getElementById("attachments-count");
    badge.textContent = attachments.length || "";
    if (!attachments.length) {
        list.innerHTML = `<div class="compose-empty">No attachments yet. Attach a document then click &#8635; Refresh.</div>`;
        document.getElementById("compose-bundle-footer").classList.add("hidden");
        document.getElementById("btn-compose-upload").disabled = true;
        return;
    }
    if(attachments.length === 1) {
        bulkSwitch.checked = false;
        bulkSwitch.disabled = true;
    } else bulkSwitch.disabled = false;

    list.innerHTML = "";
    if (isComposeBulkMode()) {
        document.getElementById("compose-bundle-footer").classList.remove("hidden");
        document.getElementById("btn-compose-upload").disabled = false;
        renderBundleList(attachments, list);
    } else {
        document.getElementById("compose-bundle-footer").classList.add("hidden");
        renderIndividualComposeList(attachments, list);
    }
}
async function handleComposeBundleUpload() {
    const bundleFooter = document.getElementById("compose-bundle-footer");
    const primaryRadio = document.querySelector("input[name='primaryIndex']:checked");
    if (!primaryRadio) { showComposeStatus("Please select a primary document."); return; }
    const primaryIndex  = parseInt(primaryRadio.value);
    const primaryAtt    = _composeAttachments[primaryIndex];
    const secondaryIndices = Array.from(document.querySelectorAll("input[name='secondaryIndex']:checked"))
        .map(c => parseInt(c.value)).filter(i => i !== primaryIndex);
    const uploadBtn = document.getElementById("btn-compose-upload");
    uploadBtn.disabled = true;
    document.getElementById("compose-result").classList.add("hidden");
    try {
        showComposeStatus("Signing in\u2026");
        const token = await getAuthToken();
        showComposeStatus("Uploading primary document\u2026");
        const { conversationId, documentId } = await uploadPrimary(primaryAtt, token);
        if (secondaryIndices.length > 0) {
            showComposeStatus("Uploading " + secondaryIndices.length + " supporting doc(s)\u2026");
            for (const idx of secondaryIndices) await uploadSupportingById(_composeAttachments[idx], conversationId, token);
        }
        showComposeStatus("Creating share link\u2026");
        await callShareApi(token, conversationId, documentId, _senderEmail, _composeRecipients);
        showComposeStatus("Inserting link into email\u2026");
        const documentURL = `${BLUE_BASE}/conversation?conversation-id=${conversationId}&doc-id=${documentId}`;
        await insertShareLinkIntoBody(documentURL, primaryAtt.name);
        const allAttIds = [primaryAtt.id, ...secondaryIndices.map(i => _composeAttachments[i].id)];
        state.suppressAttachmentRefresh = true;
        await removeAttachmentIfRequested(allAttIds);
        if (_customProps) {
            saveConversationRecord(_customProps, `compose_${conversationId}`, {
                conversationId, documentId,
                label: primaryAtt.name, uploadType: "bundle", timestamp: Date.now(),
            }).catch(err => console.warn("customProps save failed:", err.message));
        }
        document.getElementById("compose-bundle-footer").classList.add("hidden");
        showComposeStatus(""); renderComposeResult(documentURL);
    } catch (err) {
        console.error("Compose bundle upload error:", err); showComposeStatus("Error: " + err.message); clearToken();
    } finally { uploadBtn.disabled = false; }
}
async function handleComposeSingleUpload(index) {
    const att = _composeAttachments[index];
    document.querySelectorAll(".btn-upload-share").forEach(b => b.disabled = true);
    document.getElementById("compose-result").classList.add("hidden");
    showComposeStatus("Signing in\u2026");
    try {
        const token = await getAuthToken();
        showComposeStatus("Uploading " + att.name + "\u2026");
        const { conversationId, documentId } = await uploadPrimary(att, token);
        showComposeStatus("Creating share link\u2026");
        await callShareApi(token, conversationId, documentId, _senderEmail, _composeRecipients);
        showComposeStatus("Inserting link into email\u2026");
        const documentURL = `${BLUE_BASE}/conversation?conversation-id=${conversationId}&doc-id=${documentId}`;
        await insertShareLinkIntoBody(documentURL, att.name);
        state.suppressAttachmentRefresh = true;
        await removeAttachmentIfRequested([att.id]);
        if (_customProps) {
            saveConversationRecord(_customProps, `compose_${conversationId}`, {
                conversationId, documentId,
                label: att.name, uploadType: "single", timestamp: Date.now(),
            }).catch(err => console.warn("customProps save failed:", err.message));
        }
        showComposeStatus(""); 
        renderComposeResult(documentURL);
    } catch (err) {
        console.error("Compose single upload error:", err); showComposeStatus("Error: " + err.message); clearToken();
    } finally { document.querySelectorAll(".btn-upload-share").forEach(b => b.disabled = false); }
}
function copyResultLink() {
    const link = document.getElementById("result-link-text").textContent;
    const btn  = document.getElementById("btn-copy-link");
    const done = () => { btn.classList.add("copied"); setTimeout(() => btn.classList.remove("copied"), 1600); };
    if (navigator.clipboard) navigator.clipboard.writeText(link).then(done).catch(() => fallbackCopy(link, done));
    else fallbackCopy(link, done);
}

// ══════════════════════════════════════════════════════════════════════════
// CHAT
// ══════════════════════════════════════════════════════════════════════════
async function enterChat(conversationId, documentId, token) {
    state.currentConversationId = conversationId;
    state.currentDocumentId = documentId;
    document.getElementById("view-read-init").classList.add("hidden");
    document.getElementById("view-read").classList.add("hidden");
    document.getElementById("view-compose").classList.add("hidden");
    document.getElementById("view-chat").classList.remove("hidden");
    document.getElementById("btn-back").classList.remove("hidden");
    document.getElementById("chat-history").innerHTML = "";
    hideSuggestions();
    if (!documentId) {
        appendMessage("ai", "Cannot start chat: document ID is missing. Please re-share the document.", conversationId, documentId); return;
    }
    showTypingIndicator();
    // Step 1: GET /v1/conversation/{id}
    // If the conversation already has assistant messages, welcome was already
    // called once — restore history and skip welcome entirely.
    try {
        const histResp = await fetchHistory(token, conversationId);
        if (histResp.ok) {
            const conv  = await histResp.json();
            const msgs  = Array.isArray(conv.messages) ? conv.messages : [];
            const hasAI = msgs.some(m => m.sender === "assistant" || m.role === "assistant");
            if (hasAI) {
                hideTypingIndicator();
                restoreConversationHistory(msgs, conversationId, documentId); // skip welcome
                return;
            }
        }
    } catch (histErr) {
        // Non-fatal: CORS / network / 4xx
        console.warn("History fetch failed (non-fatal):", histErr.message);
    }

    // Step 2: No prior history — call welcome (first open only)
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
        try { data = JSON.parse(rawText); } catch { appendMessage("ai", "Hello! How can I help you with this document?", conversationId, documentId); return; }
        const welcomeMsg = data.answer || data.response || data.message || data.text ||
            data.content || data.welcomeText || data.welcome_text || (typeof data === "string" ? data : null);
        appendMessage("ai", welcomeMsg || "Hello! How can I help you with this document?", conversationId, documentId);
        const tags = Array.isArray(data.tags) ? data.tags : [];
        if (tags.length) renderSuggestions(tags);
    } catch (err) {
        hideTypingIndicator();
        console.error("enterChat welcome error:", err);
        appendMessage("ai", "Network error. You can still ask questions below.", conversationId, documentId);
    }
}
async function sendChatMessage() {
    const input = document.getElementById("user-input");
    const text  = input.value.trim();
    if (!text) return;
    hideSuggestions(); appendMessage("user", text, state.currentConversationId, state.currentDocumentId); input.value = "";
    document.getElementById("btn-send").disabled = true;
    showTypingIndicator();
    try {
        const token = await getAuthToken();
        const resp  = await askQuestion(token, state.currentConversationId, text);
        hideTypingIndicator();
        if (!resp.ok) throw new Error("Ask failed (" + resp.status + "): " + await resp.text());
        const data = await resp.json();
        appendMessage("ai", data.answer || data.response || "No response received.", state.currentConversationId, state.currentDocumentId);
        const tags = Array.isArray(data.tags) ? data.tags : [];
        if (tags.length) renderSuggestions(tags);
    } catch (err) {
        hideTypingIndicator(); appendMessage("ai", "Error: " + err.message, state.currentConversationId, state.currentDocumentId); clearToken();
    } finally { document.getElementById("btn-send").disabled = false; }
}