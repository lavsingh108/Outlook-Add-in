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
        authority:   "https://login.microsoftonline.com/common", // multi-tenant
        redirectUri: window.location.href.split("?")[0],
    },
    cache: { cacheLocation: "sessionStorage", storeAuthStateInCookie: false },
};

// ── State ──────────────────────────────────────────────────────────────────
const state = { currentConversationId: null, currentDocumentId: null, suppressAttachmentRefresh: false };

let _msal                 = null;
let _cachedSmartBlueToken = null;
let _customProps          = null;
let _composeAttachments        = [];
let _composeRecipients         = [];
let _senderEmail               = "";
let _readShareInfo             = null;  // share link found in email body; set in initRead
let _composeConversationCtx    = null;       // { conversationId, documentId } after first compose upload
let _composeUploadedAttIds     = new Set();  // att.id values uploaded this compose session
let _composeSharedRecipients   = new Set();  // recipient emails already shared with
let _composeRefreshTimer       = null;        // debounce timer for change events
let _chatFromCompose           = false;       // true when chat opened from compose mode
let _composeAccessLevel        = "restricted"; // "restricted" | "anonymous"

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

// Returns the stored conversation record for this attachment, or null if not uploaded.
function getAttachmentRecord(att) {
    if (!_customProps) return null;
    const map = getConversationMap(_customProps);
    if (map[singleFingerprint(att)]) return map[singleFingerprint(att)];
    const bundleKey = Object.keys(map).find(fp => fp.startsWith("bundle_") && fp.includes(att.id));
    return bundleKey ? map[bundleKey] : null;
}
function isAttachmentUploaded(att) { return !!getAttachmentRecord(att); }
function bundleFingerprint(primaryAtt, secondaryAtts) {
    const ids = secondaryAtts.map(a => a.id).sort();
    return `bundle_${[primaryAtt.id, ...ids].join("_")}`;
}

// ── Thread-scoped context (roamingSettings) ───────────────────────────────
// Office.context.roamingSettings syncs via Exchange — same account, any device,
// any Outlook client. Keyed by Outlook's conversationId which is identical for
// every email in the same thread.
//
// Key:   "sb_thread_{outlookConversationId}"
// Value: Array of { conversationId, documentId, label, uploadType, timestamp }
//
// 32KB total limit. Each record ≈ 150 bytes → ~200 threads before pruning.
// Entries older than 90 days are pruned on every write.

function getThreadKey() {
    const threadId = Office.context.mailbox.item.conversationId || "unknown";
    return `sb_thread_${threadId}`;
}

function saveThreadContext(record) {
    try {
        const rs  = Office.context.roamingSettings;
        const key = getThreadKey();

        // Merge by conversationId — no duplicates
        const existing = getThreadContextAll();
        const map = {};
        existing.forEach(r => { map[r.conversationId] = r; });
        map[record.conversationId] = record;
        rs.set(key, Object.values(map));

        // Maintain a key index so we can prune old threads
        const cutoff    = Date.now() - 90 * 24 * 60 * 60 * 1000;
        const index     = rs.get("sb_thread_index") || [];
        if (!index.includes(key)) index.push(key);
        const activeIdx = index.filter(k => {
            const recs = rs.get(k);
            if (!recs || !recs.length) { rs.remove(k); return false; }
            const latest = Math.max(...recs.map(r => r.timestamp || 0));
            if (latest < cutoff) { rs.remove(k); return false; }
            return true;
        });
        rs.set("sb_thread_index", activeIdx);

        rs.saveAsync(result => {
            if (result.status !== Office.AsyncResultStatus.Succeeded)
                console.warn("roamingSettings save failed:", result.error?.message);
            else
                console.log("Thread context saved to roamingSettings:", key);
        });
    } catch (e) { console.warn("saveThreadContext failed:", e.message); }
}

function getThreadContextAll() {
    try {
        return Office.context.roamingSettings.get(getThreadKey()) || [];
    } catch { return []; }
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

        const shareId = sp.get("share-id") || sp.get("shareId") || null;
        if (shareId) return { conversationId: null, docId: null, shareId, shareUrl: url };
        if (conversationId) return { conversationId, docId, shareUrl: url };
    } catch (_) {}
    return { conversationId: null, docId: null, shareId: null, shareUrl: null };
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
                if (parsed.conversationId || parsed.shareId) {
                    resolve({ ...parsed, linkText: linkText || m[1] }); return;
                }
            }
            const urlRe = /https?:\/\/[^\s"'<>)]+/gi;
            while ((m = urlRe.exec(html)) !== null) {
                const parsed = parseDocUrl(m[0]);
                if (parsed.conversationId || parsed.shareId) {
                    resolve({ ...parsed, linkText: null }); return;
                }
            }
            resolve({ conversationId: null, docId: null, shareId: null, linkText: null });
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
// attachmentIds: single id string OR array of ids (primary + supporting docs).
// removeAttachmentAsync accepts one id at a time — iterate and remove sequentially.
// Always resolves — user-added attachments fail silently (non-fatal).
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

async function resolveShareId(shareId, token) {
    const resp = await fetch(
        `${PROXY_BASE}/v1/conversation/share/${encodeURIComponent(shareId)}`,
        { headers: { Authorization: "Bearer " + token, "ngrok-skip-browser-warning": "true" } }
    );
    if (!resp.ok) throw new Error("Share resolve failed (" + resp.status + "): " + await resp.text());
    const data = await resp.json();
    const conversationId = data.conversation_id || data.conversationId || null;
    const docId          = data.doc_id          || data.docId          || null;
    if (!conversationId) throw new Error("No conversation_id returned from share resolve.");
    if (!docId)          throw new Error("No doc_id returned from share resolve.");
    return { conversationId, docId };
}

async function getShareLink(token, conversationId, docId, access = 'restricted') {
    // Step 1: fetch current recipient list
    const emailList = await fetch(
        `${PROXY_BASE}/v1/doc-access/share/${encodeURIComponent(docId)}/list`,
        { headers: { Authorization: "Bearer " + token, "ngrok-skip-browser-warning": "true" } }
    ).then(resp => {
        if (!resp.ok) throw new Error("Recipient list fetch failed (" + resp.status + ")");
        return resp.json();
    }).then(data => {
        const users = data.users || [];
        // emailList is a flat array of email strings
        return users.map(u => u.email).filter(Boolean);
    }).catch(err => {
        console.warn("Could not fetch recipient list:", err.message);
        return [];
    });

    // Step 2: POST to share endpoint
    const resp = await fetch(
        `${PROXY_BASE}/v1/document/${encodeURIComponent(docId)}/share`,
        {
            method: "POST",
            headers: { Authorization: "Bearer " + token, "ngrok-skip-browser-warning": "true", "Content-Type": "application/json" },
            body: JSON.stringify({
                receivers:        emailList.map(email => ({ email })),
                allowed_domains:  ["*"],
                roles_allowed:    [access],
                expire_in_secs:   30 * 24 * 60 * 60,
                allow_download:   false,
                allow_handsfree:  false,
                text_notes:       [""],
                voice_notes:      [],
            }),
        }
    );
    if (!resp.ok) throw new Error("Share link API failed (" + resp.status + "): " + await resp.text());
    const data = await resp.json();
    const url = data["share-url"] || "";
    if (!url) throw new Error("Share link API returned no URL.");
    return url;
}

function fetchHistory(token, conversationId) {
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
                    ${isAttachmentUploaded(att) && !isPrimary
                        ? '<span class="sec-label" style="color:#2e7d32">\u2713 Uploaded</span>'
                        : `<input type="checkbox" name="secondaryIndex" value="${index}" id="chk-sec-${index}"
                               ${isPrimary ? "" : "checked"} ${isPrimary ? "disabled" : ""}/>
                           <label class="sec-label" for="chk-sec-${index}">Include</label>`
                    }
                </div>
            </div>`;
        container.appendChild(div);
    });
    container.querySelectorAll("input[name='primaryIndex']").forEach(radio => {
        radio.addEventListener("change", () => {
            updateBundleSelection(container);
            // Re-evaluate footer button when primary selection changes
            const atts = Office.context.mailbox.item.attachments ||
                         _composeAttachments || [];
            const selectedIndex = parseInt(radio.value);
            const primaryAtt = atts[selectedIndex];
            if (!primaryAtt) return;
            const bundleBtn = document.getElementById("btn-upload-bundle");
            if (!bundleBtn) return;
            const rec = getAttachmentRecord(primaryAtt);
            if (rec) {
                bundleBtn.textContent = "\u25B6 Start Chat";
                bundleBtn.classList.add("btn-start-chat-att");
                bundleBtn.onclick = async () => {
                    bundleBtn.disabled = true;
                    showReadStatus("Signing in\u2026");
                    try { const token = await getAuthToken(); await enterChat(rec.conversationId, rec.documentId, token); }
                    catch (err) { showReadStatus("Error: " + err.message); clearToken(); bundleBtn.disabled = false; }
                };
            } else {
                bundleBtn.textContent = "\u2B06 Upload & Analyse";
                bundleBtn.classList.remove("btn-start-chat-att");
                bundleBtn.onclick = handleReadBundleUpload;
            }
        });
    });
}
function updateBundleSelection(container) {
    const primaryVal = container.querySelector("input[name='primaryIndex']:checked")?.value;
    container.querySelectorAll(".att-item").forEach(item => {
        const idx = item.dataset.index; const isPrimary = idx === primaryVal;
        const secChk = item.querySelector("input[name='secondaryIndex']");
        item.classList.toggle("is-primary", isPrimary);
        if (secChk) {
            if (isPrimary) { secChk.checked = false; secChk.disabled = true; }
            else { secChk.disabled = false; if (!secChk.dataset.userUnchecked) secChk.checked = true; }
        }
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
        const rec = getAttachmentRecord(att);
        // Any stored record → Start Chat (points to that conversation)
        const btnLabel = rec ? "\u25B6 Start Chat"
            : hasContext ? "Add to Bundle" : "Upload";
        const btnClass = rec ? "btn-upload-single btn-start-chat-att" : "btn-upload-single";
        div.innerHTML = `
            <div class="att-individual-row">
                <div class="att-info">
                    <div class="att-name" title="${escHtml(att.name)}">${escHtml(att.name)}</div>
                    <div class="att-meta">${formatBytes(att.size)}</div>
                </div>
                <button class="${btnClass}" data-index="${index}">${btnLabel}</button>
            </div>`;
        container.appendChild(div);
    });
    container.querySelectorAll(".btn-upload-single").forEach(btn => {
        const i = parseInt(btn.dataset.index);
        const r = getAttachmentRecord(Office.context.mailbox.item.attachments[i]);
        if (r) {
            btn.onclick = async () => {
                btn.disabled = true;
                showReadStatus("Signing in\u2026");
                try { const token = await getAuthToken(); await enterChat(r.conversationId, r.documentId, token); }
                catch (err) { showReadStatus("Error: " + err.message); clearToken(); btn.disabled = false; }
            };
        } else {
            btn.onclick = hasContext
                ? () => handleReadAddToExisting(i)
                : () => handleReadSingleUpload(i);
        }
    });
}
// Simplified list for "Add to Bundle" mode — all attachments are selectable
// as supporting docs (no primary concept since we already have a conversation).
function renderAddToBundleList(attachments, container) {
    container.innerHTML = "";
    attachments.forEach((att, index) => {
        const div = document.createElement("div");
        div.className = "att-item";
        div.dataset.index = index;
        const alreadyAdded = isAttachmentUploaded(att);
        div.innerHTML = `
            <div class="att-bundle-row">
                <div class="att-secondary-col">
                    ${alreadyAdded
                        ? '<span class="sec-label" style="color:#2e7d32">\u2713 Added</span>'
                        : `<input type="checkbox" name="addToBundleIndex" value="${index}"
                               id="chk-add-${index}" checked />
                           <label class="sec-label" for="chk-add-${index}">Include</label>`
                    }
                </div>
                <div class="att-info">
                    <div class="att-name" title="${escHtml(att.name)}">${escHtml(att.name)}</div>
                    <div class="att-meta">${formatBytes(att.size)}</div>
                </div>
            </div>`;
        container.appendChild(div);
    });
}

function renderIndividualComposeList(attachments, container) {
    attachments.forEach((att, index) => {
        const div = document.createElement("div");
        div.className = "att-item";
        // Use in-memory Set — _customProps is not loaded in compose mode
        const alreadyUploaded = _composeUploadedAttIds.has(att.id);
        const btnLabel = alreadyUploaded
            ? "\u2713 Shared"
            : _composeConversationCtx ? "Add to Bundle" : "Share";
        const btnClass = alreadyUploaded
            ? "btn-upload-single btn-upload-share btn-shared-done"
            : _composeConversationCtx
                ? "btn-upload-single btn-upload-share btn-add-bundle"
                : "btn-upload-single btn-upload-share";
        div.innerHTML = `
            <div class="att-individual-row">
                <div class="att-info">
                    <div class="att-name" title="${escHtml(att.name)}">${escHtml(att.name)}</div>
                    <div class="att-meta">${formatBytes(att.size)}</div>
                </div>
                <button class="${btnClass}" data-index="${index}" data-att-id="${att.id}"
                    ${alreadyUploaded ? 'disabled' : ''}>${btnLabel}</button>
            </div>`;
        container.appendChild(div);
    });
    container.querySelectorAll(".btn-upload-share:not([disabled])").forEach(btn => {
        const i = parseInt(btn.dataset.index);
        if (_composeConversationCtx) {
            btn.onclick = () => handleComposeAddToBundle(i);
        } else {
            btn.onclick = () => handleComposeSingleUpload(i);
        }
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
            let { conversationId, docId } = shareInfo;
            // share-id URL — resolve to conversation/doc IDs first
            if (!conversationId && shareInfo.shareId) {
                showReadStatus("Resolving share link\u2026");
                const resolved = await resolveShareId(shareInfo.shareId, token);
                conversationId = resolved.conversationId;
                docId          = resolved.docId;
                shareInfo.conversationId = conversationId;
                shareInfo.docId          = docId;
            }
            await enterChat(conversationId, docId, token);
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
    // Store URL in hidden input for copyResultLink fallback
    document.getElementById("result-link-text").value = link;
    // Copy URL button — already in HTML
    const copyBtn = document.getElementById("btn-copy-link");
    if (copyBtn) {
        copyBtn.onclick = () => {
            const restore = () => {
                setTimeout(() => {
                    copyBtn.innerHTML = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" width="14" height="14"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg> Copy URL';
                }, 1600);
            };
            const done = () => { copyBtn.textContent = "\u2713 Copied!"; restore(); };
            if (navigator.clipboard) navigator.clipboard.writeText(link).then(done).catch(() => fallbackCopy(link, done));
            else fallbackCopy(link, done);
        };
    }
    // Curate button — already in HTML
    const curateBtn = document.getElementById("btn-curate-link");
    if (curateBtn) curateBtn.onclick = () => window.open(link, "_blank");
    const startChatBtn = document.getElementById("btn-compose-start-chat");
    if (startChatBtn) {
        startChatBtn.classList.remove("hidden");
        startChatBtn.onclick = async () => {
            if (!_composeConversationCtx) return;
            startChatBtn.disabled = true;
            try {
                const token = await getAuthToken();
                _chatFromCompose = true;
                await enterChat(_composeConversationCtx.conversationId, _composeConversationCtx.documentId, token);
            } catch (err) {
                showComposeStatus("Error: " + err.message); clearToken();
                startChatBtn.disabled = false;
            }
        };
    }
    document.getElementById("compose-access-row")?.classList.add("hidden");
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
            if (shareInfo && (shareInfo.conversationId || shareInfo.shareId)) {
                _readShareInfo = shareInfo;  // persist for attachment rendering
                renderShareSection(shareInfo);
            }

            loadCustomProps()
                .then(cp => {
                    _customProps = cp;

                    // Always render Previous Chats and attachments first —
                    // this ensures ALL thread emails show the shared context
                    // (via the synthetic share-link record in renderPreviousChats).
                    renderPreviousChats();
                    loadReadAttachments();
                    const atts = Office.context.mailbox.item.attachments || [];
                    if (shareInfo && shareInfo.conversationId && atts.length > 0)
                        document.getElementById("read-or-divider").classList.remove("hidden");

                    // Auto-open: only resume a stored primary record (not share-link only).
                    // Share link has its own Start Chat button — no silent auto-open.
                    const cpPrimary = Object.values(getConversationMap(cp))
                        .filter(r => r.uploadType !== "shared-link");
                    const thPrimary = getThreadContextAll().filter(r => r.uploadType !== "shared-link");
                    const primaryRecs = [...cpPrimary,
                        ...thPrimary.filter(r => !cpPrimary.some(c => c.conversationId === r.conversationId))]
                        .sort((a, b) => (b.timestamp || 0) - (a.timestamp || 0));

                    if (primaryRecs.length > 0) {
                        const latest = primaryRecs[0];
                        getAuthToken()
                            .then(token => enterChat(latest.conversationId, latest.documentId, token))
                            .catch(err => console.warn("Auto-resume failed:", err.message));
                    } else if (shareInfo && shareInfo.shareId) {
                        getAuthToken()
                            .then(async token => {
                                const { conversationId, docId } = await resolveShareId(shareInfo.shareId, token);
                                shareInfo.conversationId = conversationId;
                                shareInfo.docId          = docId;
                                await enterChat(conversationId, docId, token);
                            })
                            .catch(err => console.warn("Share-id auto-open failed:", err.message));
                    } else if (shareInfo && shareInfo.conversationId) {
                        getAuthToken()
                            .then(token => enterChat(shareInfo.conversationId, shareInfo.docId, token))
                            .catch(err => console.warn("Share link auto-open failed:", err.message));
                    }
                })
                .catch(() => {
                    // Custom props unavailable — still render what we can from share link
                    renderPreviousChats();
                    loadReadAttachments();
                    const atts = Office.context.mailbox.item.attachments || [];
                    if (shareInfo && shareInfo.conversationId && atts.length > 0)
                        document.getElementById("read-or-divider").classList.remove("hidden");
                });
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
    // Merge custom props records with thread-scoped sessionStorage records.
    const storedMap = {};
    Object.values(getConversationMap(_customProps))
        .filter(r => r.uploadType !== "shared-link")
        .forEach(r => { storedMap[r.conversationId] = r; });
    getThreadContextAll()
        .filter(r => r.uploadType !== "shared-link")
        .forEach(r => { if (!storedMap[r.conversationId]) storedMap[r.conversationId] = r; });
    const records = Object.values(storedMap);
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
    // hasContext: custom props OR thread sessionStorage. Share link is NEVER a bundle target.
    const cpRecords = _customProps
        ? Object.values(getConversationMap(_customProps)).filter(r => r.uploadType !== "shared-link")
        : [];
    const threadRecords = getThreadContextAll().filter(r => r.uploadType !== "shared-link");
    const hasContext = cpRecords.length > 0 || threadRecords.length > 0;

    if (isReadBulkMode()) {
        footerDiv.classList.remove("hidden");
        const bundleBtn = document.getElementById("btn-upload-bundle");
        bundleBtn.disabled = false;
        if (hasContext) {
            renderAddToBundleList(attachments, listDiv);  // checkboxes only, no radio
            // If every attachment is already added, disable the footer button
            const allAdded = attachments.every(a => isAttachmentUploaded(a));
            if (allAdded) {
                bundleBtn.textContent = "\u2713 All Added";
                bundleBtn.disabled = true;
                bundleBtn.classList.remove("btn-start-chat-att");
                bundleBtn.onclick = null;
            } else {
                bundleBtn.textContent = "\uFF0B Add to Bundle";
                bundleBtn.disabled = false;
                bundleBtn.classList.remove("btn-start-chat-att");
                bundleBtn.onclick = handleReadAddToBundle;
            }
        } else {
            renderBundleList(attachments, listDiv);       // radio + checkboxes
            // After rendering, check if the default primary (index 0) has a record
            const primaryRec = getAttachmentRecord(attachments[0]);
            if (primaryRec) {
                bundleBtn.textContent = "\u25B6 Start Chat";
                bundleBtn.classList.add("btn-start-chat-att");
                bundleBtn.onclick = async () => {
                    bundleBtn.disabled = true;
                    showReadStatus("Signing in\u2026");
                    try { const token = await getAuthToken(); await enterChat(primaryRec.conversationId, primaryRec.documentId, token); }
                    catch (err) { showReadStatus("Error: " + err.message); clearToken(); bundleBtn.disabled = false; }
                };
            } else {
                bundleBtn.textContent = "\u2B06 Upload & Analyse";
                bundleBtn.classList.remove("btn-start-chat-att");
                bundleBtn.onclick = handleReadBundleUpload;
            }
        }
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
        saveThreadContext({ conversationId, documentId, label: primaryAtt.name, uploadType: "bundle", timestamp: Date.now() });
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
        saveThreadContext({ conversationId, documentId, label: att.name, uploadType: "single", timestamp: Date.now() });
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
    const secondaryAtts = Array.from(document.querySelectorAll("input[name='addToBundleIndex']:checked"))
        .map(c => parseInt(c.value)).map(i => attachments[i]);
    if (!secondaryAtts.length) { showReadStatus("Select at least one document to add."); return; }

    const cpRecs = Object.values(getConversationMap(_customProps || {}))
        .filter(r => r.uploadType !== "shared-link");
    const thRecs = getThreadContextAll().filter(r => r.uploadType !== "shared-link");
    const allRecs = [...cpRecs, ...thRecs.filter(r => !cpRecs.some(c => c.conversationId === r.conversationId))];
    const latestRecord = allRecs.sort((a, b) => (b.timestamp || 0) - (a.timestamp || 0))[0];
    const existingConvId = latestRecord?.conversationId;
    const existingDocId  = latestRecord?.documentId;
    if (!existingConvId) { showReadStatus("No existing conversation found."); return; }

    document.getElementById("btn-upload-bundle").disabled = true;
    showReadStatus("Signing in\u2026");
    try {
        const token = await getAuthToken();
        showReadStatus("Adding " + secondaryAtts.length + " doc(s) to conversation\u2026");
        for (const att of secondaryAtts) {
            await uploadSupportingById(att, existingConvId, token);
        }
        // Persist each added attachment so it shows as "✓ Added" on re-open
        if (_customProps) {
            await Promise.all(secondaryAtts.map(att =>
                saveConversationRecord(_customProps, singleFingerprint(att), {
                    conversationId: existingConvId,
                    documentId: existingDocId,
                    label: att.name,
                    uploadType: "bundle",
                    timestamp: Date.now(),
                }).catch(err => console.warn("customProps save failed:", err.message))
            ));
        }
        showReadStatus("");
        await enterChat(existingConvId, existingDocId, token);
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
        const cpR = Object.values(getConversationMap(_customProps || {}))
            .filter(r => r.uploadType !== "shared-link");
        const thR = getThreadContextAll().filter(r => r.uploadType !== "shared-link");
        const allR = [...cpR, ...thR.filter(r => !cpR.some(c => c.conversationId === r.conversationId))];
        const latestRec = allR.sort((a, b) => (b.timestamp || 0) - (a.timestamp || 0))[0];
        const existingConvId = latestRec?.conversationId;
        const existingDocId  = latestRec?.documentId;
        if (!existingConvId) throw new Error("No existing conversation found.");
        showReadStatus("Adding " + att.name + " to conversation\u2026");
        await uploadSupportingById(att, existingConvId, token);
        // Persist so it shows as "✓ Added" on re-open
        if (_customProps) {
            saveConversationRecord(_customProps, singleFingerprint(att), {
                conversationId: existingConvId,
                documentId: existingDocId,
                label: att.name,
                uploadType: "bundle",
                timestamp: Date.now(),
            }).catch(err => console.warn("customProps save failed:", err.message));
        }
        showReadStatus("");
        await enterChat(existingConvId, existingDocId, token);
    } catch (err) {
        console.error("Add to existing error:", err); showReadStatus("Error: " + err.message); clearToken();
        document.querySelectorAll(".btn-upload-single").forEach(b => b.disabled = false);
    }
}

function switchToReadView() {
    document.getElementById("view-chat").classList.add("hidden");
    // Return to compose if chat was opened from compose mode
    if (_chatFromCompose) {
        _chatFromCompose = false;
        document.getElementById("view-compose").classList.remove("hidden");
        document.getElementById("btn-back").classList.add("hidden");
        document.getElementById("chat-history").innerHTML = "";
        hideSuggestions();
        state.currentConversationId = null; state.currentDocumentId = null;
        const startChatBtn = document.getElementById("btn-compose-start-chat");
        if (startChatBtn) startChatBtn.disabled = false;
        return;
    }
    document.getElementById("view-read").classList.remove("hidden");
    document.getElementById("btn-back").classList.add("hidden");
    document.getElementById("chat-history").innerHTML = "";
    hideSuggestions();
    state.currentConversationId = null; state.currentDocumentId = null;
    // _readShareInfo intentionally kept — email body hasn't changed,
    // share link must stay available for hasContext and Add to Bundle after Back
    const shareBtn = document.getElementById("btn-share-chat");
    if (shareBtn) shareBtn.disabled = false;
    renderPreviousChats(); loadReadAttachments(); showReadStatus("");
}

// ══════════════════════════════════════════════════════════════════════════
// COMPOSE MODE
// ══════════════════════════════════════════════════════════════════════════
function initCompose() {
    document.querySelector(".header-title").textContent = "Share Document";
    document.getElementById("view-compose").classList.remove("hidden");
    document.getElementById("btn-refresh").classList.remove("hidden");
    document.getElementById("btn-refresh").onclick        = () => {
        clearTimeout(_composeRefreshTimer);
        state.suppressAttachmentRefresh = false;
        _composeConversationCtx  = null;
        _composeUploadedAttIds   = new Set();
        _composeSharedRecipients = new Set();
        // Restore UI hidden during post-upload state
        document.getElementById("clbl-bundle")?.classList.remove("hidden");
        document.getElementById("clbl-individual")?.classList.remove("hidden");
        document.getElementById("chk-compose-bulk").closest("label")?.classList.remove("hidden");
        document.getElementById("compose-new-recipients")?.remove();
        document.getElementById("compose-documents-section")?.classList.remove("hidden");
        document.getElementById("compose-attachment-option")?.classList.remove("hidden");
        document.getElementById("btn-compose-start-chat")?.classList.add("hidden");
        document.getElementById("compose-access-row")?.classList.remove("hidden");
        loadComposeData(true);
    };
    document.getElementById("btn-back").onclick           = switchToReadView;
    document.getElementById("sel-access").value           = _composeAccessLevel;
    document.getElementById("sel-access").onchange = () => {
        _composeAccessLevel = document.getElementById("sel-access").value;
    };
    document.getElementById("btn-compose-upload").onclick = handleComposeBundleUpload;
    document.getElementById("btn-copy-link").onclick      = copyResultLink;
    document.getElementById("chk-compose-bulk").onchange  = onComposeToggleMode;
    loadComposeData(false);

    // Live sync — fire loadComposeData whenever the user adds/removes
    // an attachment or changes recipients, removing the need to manually refresh.
    // Requires Mailbox 1.8 (AttachmentsChanged) / 1.7 (RecipientsChanged).
    // The refresh button stays as a fallback for older clients.
    // Debounce — Outlook fires these events multiple times for a single user
    // action (e.g. adding one recipient triggers 3–4 RecipientsChanged events).
    // Wait 400ms after the last event before refreshing.
    const debouncedRefresh = () => {
        if (state.suppressAttachmentRefresh) return;
        clearTimeout(_composeRefreshTimer);
        _composeRefreshTimer = setTimeout(() => loadComposeData(true), 400);
    };
    if (Office.context.requirements.isSetSupported("Mailbox", "1.8")) {
        Office.context.mailbox.item.addHandlerAsync(
            Office.EventType.AttachmentsChanged, debouncedRefresh
        );
    }
    if (Office.context.requirements.isSetSupported("Mailbox", "1.7")) {
        Office.context.mailbox.item.addHandlerAsync(
            Office.EventType.RecipientsChanged, debouncedRefresh
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
    if (state.suppressAttachmentRefresh) {
        // Upload just completed — ignore event-driven refreshes until user clicks Refresh
        document.getElementById("btn-refresh").classList.remove("spinning");
        return;
    }
    if (isRefresh) {
        document.getElementById("btn-refresh").classList.add("spinning");
        // Only hide result card if no upload has happened yet this session
        if (!_composeConversationCtx) {
            document.getElementById("compose-result").classList.add("hidden");
        }
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
    const list       = document.getElementById("compose-attachments");
    const badge      = document.getElementById("attachments-count");
    const bulkSwitch = document.getElementById("chk-compose-bulk");
    badge.textContent = attachments.length || "";
    if (!attachments.length) {
        list.innerHTML = `<div class="compose-empty">No attachments yet. Attach a document then click &#8635; Refresh.</div>`;
        document.getElementById("compose-bundle-footer").classList.add("hidden");
        document.getElementById("compose-attachment-option").classList.add("hidden");
        document.getElementById("btn-compose-upload").disabled = true;
        // Still check for new recipients even with no attachments
        if (_composeConversationCtx) renderPostUploadActions();
        return;
    }
    // Show checkbox whenever there are attachments
    document.getElementById("compose-attachment-option").classList.remove("hidden");
    // Disable the toggle when there is only one attachment — bundle makes no sense
    if(attachments.length === 1) {
        bulkSwitch.checked = false;
        bulkSwitch.disabled = true;
    } else bulkSwitch.disabled = false;

    list.innerHTML = "";

    // After an upload has completed, always use individual rendering.
    // This lets each attachment show its own state independently:
    //   already uploaded → "✓ Shared" (disabled)
    //   new attachment   → "Add to Bundle"
    // The bundle toggle is hidden so the user cannot switch back to bundle UI.
    if (_composeConversationCtx) {
        document.getElementById("compose-bundle-footer").classList.add("hidden");
        document.getElementById("chk-compose-bulk").closest("label")?.classList.add("hidden");
        document.getElementById("clbl-bundle")?.classList.add("hidden");
        document.getElementById("clbl-individual")?.classList.add("hidden");
        // Only show new (not yet uploaded) attachments
        const newAtts = attachments.filter(a => !_composeUploadedAttIds.has(a.id));
        const docsSection = document.getElementById("compose-documents-section");
        if (newAtts.length === 0) {
            // All uploaded — hide the entire Documents section and checkbox
            if (docsSection) docsSection.classList.add("hidden");
            document.getElementById("compose-attachment-option").classList.add("hidden");
        } else {
            // Show section with only the new attachments
            if (docsSection) docsSection.classList.remove("hidden");
            badge.textContent = newAtts.length;
            renderIndividualComposeList(newAtts, list);
        }
        renderPostUploadActions();
        return;
    }

    // Normal pre-upload rendering
    if (isComposeBulkMode()) {
        document.getElementById("compose-bundle-footer").classList.remove("hidden");
        document.getElementById("btn-compose-upload").disabled = false;
        renderBundleList(attachments, list);
    } else {
        document.getElementById("compose-bundle-footer").classList.add("hidden");
        renderIndividualComposeList(attachments, list);
    }
}

// Shown after a successful upload. Detects:
//   New recipients (not yet shared with) → "Share with New Recipients" button
//   New attachments are already handled by renderIndividualComposeList /
//   renderBundleList via _composeUploadedAttIds + _composeConversationCtx
function renderPostUploadActions() {
    const newRecipients = _composeRecipients.filter(e => !_composeSharedRecipients.has(e));
    let section = document.getElementById("compose-new-recipients");
    if (newRecipients.length > 0) {
        if (!section) {
            section = document.createElement("div");
            section.id = "compose-new-recipients";
            section.className = "compose-section";
            document.getElementById("compose-status").insertAdjacentElement("beforebegin", section);
        }
        section.innerHTML = `
            <div class="compose-section-header">
                <span class="compose-section-title">New Recipients</span>
                <span class="compose-badge">${newRecipients.length}</span>
            </div>
            <div class="recipient-chips" style="flex-wrap:wrap;gap:4px;margin-bottom:8px">
                ${newRecipients.map(e => `<span class="recipient-chip">${escHtml(e)}</span>`).join("")}
            </div>
            <button class="btn-primary btn-share-new-recip" style="width:100%">
                Share with New Recipients
            </button>`;
        section.querySelector(".btn-share-new-recip").onclick = async () => {
            const btn = section.querySelector(".btn-share-new-recip");
            btn.disabled = true; btn.textContent = "Sharing\u2026";
            try {
                const token = await getAuthToken();
                await callShareApi(token, _composeConversationCtx.conversationId,
                    _composeConversationCtx.documentId, _senderEmail, newRecipients);
                newRecipients.forEach(e => _composeSharedRecipients.add(e));
                // Hide section immediately, flash status for 3s
                section.remove();
                showComposeStatus("\u2713 Shared with new recipients");
                setTimeout(() => showComposeStatus(""), 3000);
            } catch (err) {
                showComposeStatus("Share failed: " + err.message); clearToken();
                btn.disabled = false; btn.textContent = "Share with New Recipients";
            }
        };
    } else if (section) {
        section.remove(); // no new recipients — hide the section
    }
}

async function handleComposeBundleUpload() {
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
        const accessLevel = document.getElementById("sel-access")?.value || _composeAccessLevel;
        const shareLink = await getShareLink(token, conversationId, documentId, accessLevel);
        await insertShareLinkIntoBody(shareLink, primaryAtt.name);
        const allAttIds = [primaryAtt.id, ...secondaryIndices.map(i => _composeAttachments[i].id)];
        // Store session state so post-upload rendering knows what was uploaded
        _composeConversationCtx = { conversationId, documentId };
        allAttIds.forEach(id => _composeUploadedAttIds.add(id));
        _composeRecipients.forEach(e => _composeSharedRecipients.add(e));
        state.suppressAttachmentRefresh = true;
        await removeAttachmentIfRequested(allAttIds);
        // Clear flag after 1.5s then force refresh — picks up anything added
        // during the suppression window (new attachments / recipients)
        setTimeout(() => { state.suppressAttachmentRefresh = false; loadComposeData(true); }, 1500);
        if (_customProps) {
            saveConversationRecord(_customProps, `compose_${conversationId}`, {
                conversationId, documentId,
                label: primaryAtt.name, uploadType: "bundle", timestamp: Date.now(),
            }).catch(err => console.warn("customProps save failed:", err.message));
        }
        saveThreadContext({ conversationId, documentId, label: primaryAtt.name, uploadType: "bundle", timestamp: Date.now() });
        document.getElementById("compose-bundle-footer").classList.add("hidden");
        showComposeStatus(""); 
        renderComposeResult(shareLink);
    } catch (err) {
        console.error("Compose bundle upload error:", err); showComposeStatus("Error: " + err.message); clearToken();
    } finally { uploadBtn.disabled = false; }
}
// Add a compose attachment as supporting doc to the already-uploaded conversation
async function handleComposeAddToBundle(index) {
    if (!_composeConversationCtx) return;
    // Use data-att-id to find the correct attachment regardless of filtered index
    const btn = document.querySelector(`.btn-upload-share[data-index="${index}"]`);
    const attId = btn?.dataset.attId;
    const att = attId
        ? _composeAttachments.find(a => a.id === attId)
        : _composeAttachments[index];
    if (!att) return;
    document.querySelectorAll(".btn-upload-share").forEach(b => b.disabled = true);
    showComposeStatus("Adding " + att.name + " to bundle\u2026");
    try {
        const token = await getAuthToken();
        await uploadSupportingById(att, _composeConversationCtx.conversationId, token);
        if (_customProps) {
            saveConversationRecord(_customProps, singleFingerprint(att), {
                conversationId: _composeConversationCtx.conversationId,
                documentId: _composeConversationCtx.documentId,
                label: att.name, uploadType: "bundle", timestamp: Date.now(),
            }).catch(err => console.warn("customProps save failed:", err.message));
        }
        _composeUploadedAttIds.add(att.id);
        // Remove all attachments from email if checkbox is checked
        const allAttIds = _composeAttachments.map(a => a.id);
        state.suppressAttachmentRefresh = true;
        await removeAttachmentIfRequested(allAttIds);
        setTimeout(() => { state.suppressAttachmentRefresh = false; }, 1500);
        // Hide the attachment row and the whole Documents section immediately
        const uploadBtn = btn || document.querySelector(`.btn-upload-share[data-index="${index}"]`);
        if (uploadBtn) uploadBtn.closest(".att-item")?.remove();
        document.getElementById("compose-documents-section")?.classList.add("hidden");
        document.getElementById("compose-attachment-option")?.classList.add("hidden");
        // Flash status for 3s then clear
        showComposeStatus("\u2713 Added to bundle");
        setTimeout(() => showComposeStatus(""), 3000);
    } catch (err) {
        console.error("Compose add to bundle error:", err);
        showComposeStatus("Error: " + err.message); clearToken();
        document.querySelectorAll(".btn-upload-share").forEach(b => b.disabled = false);
    }
}

async function handleComposeSingleUpload(index) {
    const att = _composeAttachments[index];
    // Disable all buttons while uploading to prevent double-submit
    document.querySelectorAll(".btn-upload-share").forEach(b => b.disabled = true);
    document.getElementById("compose-result").classList.add("hidden");
    showComposeStatus("Signing in\u2026");
    let succeeded = false;
    try {
        const token = await getAuthToken();
        showComposeStatus("Uploading " + att.name + "\u2026");
        const { conversationId, documentId } = await uploadPrimary(att, token);
        showComposeStatus("Creating share link\u2026");
        await callShareApi(token, conversationId, documentId, _senderEmail, _composeRecipients);
        showComposeStatus("Inserting link into email\u2026");
        const documentURL = `${BLUE_BASE}/conversation?conversation-id=${conversationId}&doc-id=${documentId}`;
        const accessLevel = document.getElementById("sel-access")?.value || _composeAccessLevel;
        const shareLink = await getShareLink(token, conversationId, documentId, accessLevel);
        await insertShareLinkIntoBody(shareLink, att.name);
        state.suppressAttachmentRefresh = true;
        await removeAttachmentIfRequested([att.id]);
        // Clear flag after short delay then force refresh
        setTimeout(() => { state.suppressAttachmentRefresh = false; loadComposeData(true); }, 1500);
        // do NOT clear flag here — AttachmentsChanged fires async after this resolves
        if (_customProps) {
            saveConversationRecord(_customProps, `compose_${conversationId}`, {
                conversationId, documentId,
                label: att.name, uploadType: "single", timestamp: Date.now(),
            }).catch(err => console.warn("customProps save failed:", err.message));
        }
        saveThreadContext({ conversationId, documentId, label: att.name, uploadType: "single", timestamp: Date.now() });
        // Store context so remaining attachments can "Add to Bundle"
        _composeConversationCtx = { conversationId, documentId };
        _composeUploadedAttIds.add(att.id);
        _composeRecipients.forEach(e => _composeSharedRecipients.add(e));
        succeeded = true;
        showComposeStatus("");
        renderComposeResult(shareLink);
        // Re-render list so remaining buttons switch to "Add to Bundle"
        renderComposeAttachments(_composeAttachments);
    } catch (err) {
        console.error("Compose single upload error:", err); showComposeStatus("Error: " + err.message); clearToken();
    } finally {
        // On success: hide the button for the uploaded doc, re-enable all others
        // On failure: re-enable all buttons so the user can retry
        document.querySelectorAll(".btn-upload-share").forEach(b => {
            const btnIndex = parseInt(b.dataset.index);
            if (succeeded && btnIndex === index) {
                b.textContent = "\u2713 Shared";   // ✓ Shared
                b.disabled = true;
                b.classList.add("btn-shared-done");
            } else {
                b.disabled = false;
            }
        });
    }
}
function copyResultLink() {
    const link = document.getElementById("result-link-text")?.value || "";
    if (link) {
        if (navigator.clipboard) navigator.clipboard.writeText(link).catch(() => fallbackCopy(link, () => {}));
        else fallbackCopy(link, () => {});
    }
}

// ══════════════════════════════════════════════════════════════════════════
// CHAT
// ══════════════════════════════════════════════════════════════════════════
async function enterChat(conversationId, documentId, token) {
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