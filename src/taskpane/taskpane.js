// ── Proxy base URL ─────────────────────────────────────────────────────────
const PROXY_BASE     = "https://headphone-crust-stipulate.ngrok-free.dev";

const AUTH_URL       = `${PROXY_BASE}/v1/authenticate`;
const UPLOAD_URL     = `${PROXY_BASE}/v1/document/upload`;
const BUNDLE_ADD_URL = `${PROXY_BASE}/v1/document/bundle/add`;
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
    if (isCompose) { initCompose(); } else { initRead(); }
}

// ══════════════════════════════════════════════════════════════════════════
// READ MODE
// ══════════════════════════════════════════════════════════════════════════

function initRead() {
    document.querySelector(".header-title").textContent = "View Document";

    document.getElementById("btn-send").onclick = sendChatMessage;
    document.getElementById("user-input").addEventListener("keydown", (e) => {
        if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); sendChatMessage(); }
    });
    document.getElementById("btn-back").onclick  = switchToReadView;
    document.getElementById("btn-upload-bundle").onclick = handleReadBundleUpload;
    document.getElementById("chk-bulk").onchange = onReadToggleMode;

    document.getElementById("view-read-init").classList.remove("hidden");

    extractShareLinkFromBody()
        .then(shareInfo => {
            document.getElementById("view-read-init").classList.add("hidden");
            document.getElementById("view-read").classList.remove("hidden");

            if (shareInfo && shareInfo.conversationId) {
                renderShareSection(shareInfo);
            }

            loadReadAttachments();

            const atts = Office.context.mailbox.item.attachments || [];
            if (shareInfo && shareInfo.conversationId && atts.length > 0) {
                document.getElementById("read-or-divider").classList.remove("hidden");
            }
        })
        .catch(err => showReadInitError("Error reading email: " + err.message));
}

// ── Parse email body for the first document URL ────────────────────────────
// Inserted URL format: ${PROXY_BASE}/conversation?conversation-id=X&doc-id=Y
function extractShareLinkFromBody() {
    return new Promise((resolve, reject) => {
        Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, (result) => {
            if (result.status !== Office.AsyncResultStatus.Succeeded) {
                reject(new Error(result.error?.message || "Body read failed"));
                return;
            }

            const html = result.value || "";

            // 1. Anchor hrefs with visible text (for display name)
            const anchorRe = /<a[^>]+href=["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>/gi;
            let m;
            while ((m = anchorRe.exec(html)) !== null) {
                const href     = m[1];
                const linkText = m[2].replace(/<[^>]+>/g, "").replace(/\s+/g, " ").trim();
                const parsed   = parseDocUrl(href);
                if (parsed.conversationId) {
                    resolve({ ...parsed, linkText: linkText || href });
                    return;
                }
            }

            // 2. Fallback: bare URLs
            const urlRe = /https?:\/\/[^\s"'<>)]+/gi;
            while ((m = urlRe.exec(html)) !== null) {
                const parsed = parseDocUrl(m[0]);
                if (parsed.conversationId) {
                    resolve({ ...parsed, linkText: null });
                    return;
                }
            }

            resolve({ conversationId: null, docId: null, linkText: null });
        });
    });
}

// Parses URLs of the form:
//   ?conversation-id=xxx&doc-id=yyy   (inserted by this add-in)
//   ?conversation_id=xxx&doc_id=yyy
//   ?conversationId=xxx&documentId=yyy
function parseDocUrl(rawUrl) {
    try {
        const url = rawUrl.replace(/[>)"'\s]+$/, "");
        const u   = new URL(url);
        const sp  = u.searchParams;

        const conversationId =
            sp.get("conversation-id") ||   // hyphenated  (proxy URL format)
            sp.get("conversation_id") ||
            sp.get("conversationId")  ||
            sp.get("cid")             || null;

        const docId =
            sp.get("doc-id")     ||         // hyphenated  (proxy URL format)
            sp.get("doc_id")     ||
            sp.get("documentId") ||
            sp.get("did")        || null;

        if (conversationId) return { conversationId, docId, shareUrl: url };

        // Path-based fallback
        const segments = u.pathname.split("/").filter(Boolean);
        if (segments.length >= 2)
            return { conversationId: segments[segments.length - 2], docId: segments[segments.length - 1], shareUrl: url };
        if (segments.length === 1)
            return { conversationId: segments[0], docId: null, shareUrl: url };
    } catch (_) {}
    return { conversationId: null, docId: null, shareUrl: null };
}

// ── Section A: Share card ─────────────────────────────────────────────────
function renderShareSection(shareInfo) {
    const section = document.getElementById("read-share-section");
    const card    = document.getElementById("read-share-card");

    const displayText = shareInfo.linkText || shareInfo.shareUrl || "View on SmartBlue";
    const displayUrl  = shareInfo.shareUrl  || "";

    card.innerHTML = `
        <div class="read-share-inner">
            <svg class="read-share-file-icon" viewBox="0 0 24 24" fill="none"
                 stroke="currentColor" stroke-width="2"
                 stroke-linecap="round" stroke-linejoin="round">
                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                <polyline points="14 2 14 8 20 8"/>
            </svg>
            <div class="read-share-info">
                <div class="read-share-name" title="${escHtml(displayText)}">${escHtml(displayText)}</div>
                <div class="read-share-url" title="${escHtml(displayUrl)}">${escHtml(displayUrl)}</div>
            </div>
        </div>
        <button class="btn-start-chat" id="btn-share-chat">
            <svg width="12" height="12" viewBox="0 0 24 24" fill="currentColor">
                <polygon points="5 3 19 12 5 21 5 3"/>
            </svg>
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
        } catch (err) {
            showReadStatus("Error: " + err.message);
            _cachedSmartBlueToken = null;
            btn.disabled = false;
        }
    };
}

// ── Section B: Attachments ─────────────────────────────────────────────────
function isReadBulkMode() { return document.getElementById("chk-bulk").checked; }

function onReadToggleMode() {
    const bulk = isReadBulkMode();
    document.getElementById("lbl-bundle").classList.toggle("active", bulk);
    document.getElementById("lbl-individual").classList.toggle("active", !bulk);
    document.getElementById("bundle-footer").classList.toggle("hidden", !bulk);
    loadReadAttachments();
}

function loadReadAttachments() {
    const attachments = Office.context.mailbox.item.attachments || [];
    const listDiv     = document.getElementById("attachment-list");
    const footerDiv   = document.getElementById("bundle-footer");

    if (!attachments || attachments.length === 0) {
        listDiv.innerHTML = "<p class='att-empty'>No attachments found.</p>";
        footerDiv.classList.add("hidden");
        document.getElementById("btn-upload-bundle").disabled = true;
        return;
    }

    if (isReadBulkMode()) {
        footerDiv.classList.remove("hidden");
        document.getElementById("btn-upload-bundle").disabled = false;
        renderBundleList(attachments, listDiv);
    } else {
        footerDiv.classList.add("hidden");
        renderIndividualReadList(attachments, listDiv);
    }
}

async function handleReadBundleUpload() {
    const attachments  = Office.context.mailbox.item.attachments;
    const primaryRadio = document.querySelector("input[name='primaryIndex']:checked");
    if (!primaryRadio) { showReadStatus("Please select a primary document."); return; }

    const primaryIndex     = parseInt(primaryRadio.value);
    const primaryAtt       = attachments[primaryIndex];
    const secondaryIndices = Array.from(
        document.querySelectorAll("input[name='secondaryIndex']:checked")
    ).map(c => parseInt(c.value)).filter(i => i !== primaryIndex);

    document.getElementById("btn-upload-bundle").disabled = true;
    showReadStatus("Signing in\u2026");

    try {
        const token = await getAuthToken();
        showReadStatus("Uploading primary document\u2026");
        const { conversationId, documentId } = await uploadPrimary(primaryAtt, token);

        if (secondaryIndices.length > 0) {
            showReadStatus("Uploading " + secondaryIndices.length + " supporting doc(s)\u2026");
            for (const idx of secondaryIndices) {
                await uploadSupporting(attachments[idx], conversationId, token);
            }
        }

        await enterChat(conversationId, documentId, token);
    } catch (err) {
        console.error("Read bundle upload error:", err);
        showReadStatus("Error: " + err.message);
        _cachedSmartBlueToken = null;
        document.getElementById("btn-upload-bundle").disabled = false;
    }
}

async function handleReadSingleUpload(index) {
    const att = Office.context.mailbox.item.attachments[index];
    document.querySelectorAll(".btn-upload-single").forEach(b => b.disabled = true);
    showReadStatus("Uploading " + att.name + "\u2026");

    try {
        const token = await getAuthToken();
        const { conversationId, documentId } = await uploadPrimary(att, token);
        await enterChat(conversationId, documentId, token);
    } catch (err) {
        console.error("Read single upload error:", err);
        showReadStatus("Error: " + err.message);
        _cachedSmartBlueToken = null;
        document.querySelectorAll(".btn-upload-single").forEach(b => b.disabled = false);
    }
}

function renderIndividualReadList(attachments, container) {
    container.innerHTML = "";
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
        btn.onclick = () => handleReadSingleUpload(parseInt(btn.dataset.index));
    });
}

function switchToReadView() {
    document.getElementById("view-chat").classList.add("hidden");
    document.getElementById("view-read").classList.remove("hidden");
    document.getElementById("btn-back").classList.add("hidden");
    document.getElementById("chat-history").innerHTML = "";
    hideSuggestions();
    currentConversationId = null;
    currentDocumentId     = null;
    const shareBtn = document.getElementById("btn-share-chat");
    if (shareBtn) shareBtn.disabled = false;
    loadReadAttachments();
    showReadStatus("");
}

function showReadStatus(msg) {
    const el = document.getElementById("status-msg");
    if (el) el.innerText = msg;
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

let _composeAttachments = [];
let _composeRecipients  = [];
let _senderEmail        = "";

function initCompose() {
    document.querySelector(".header-title").textContent = "Share Document";
    document.getElementById("view-compose").classList.remove("hidden");
    document.getElementById("btn-refresh").classList.remove("hidden");
    document.getElementById("btn-refresh").onclick = () => loadComposeData(true);
    document.getElementById("btn-compose-upload").onclick = handleComposeBundleUpload;
    document.getElementById("btn-copy-link").onclick      = copyResultLink;
    document.getElementById("chk-compose-bulk").onchange  = onComposeToggleMode;
    loadComposeData(false);
}

function isComposeBulkMode() { return document.getElementById("chk-compose-bulk").checked; }

function onComposeToggleMode() {
    const bulk = isComposeBulkMode();
    document.getElementById("clbl-bundle").classList.toggle("active", bulk);
    document.getElementById("clbl-individual").classList.toggle("active", !bulk);
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
        new Promise(res => item.to.getAsync(r =>
            res(r.status === Office.AsyncResultStatus.Succeeded ? r.value : []))),
        new Promise(res => item.cc.getAsync(r =>
            res(r.status === Office.AsyncResultStatus.Succeeded ? r.value : []))),
        new Promise(res => item.getAttachmentsAsync(r =>
            res(r.status === Office.AsyncResultStatus.Succeeded ? r.value : [])))
    ]).then(([toList, ccList, attachments]) => {
        const seen = new Set();
        _composeRecipients = [...toList, ...ccList]
            .map(r => (r.emailAddress || "").toLowerCase().trim())
            .filter(e => { if (!e || seen.has(e)) return false; seen.add(e); return true; });

        renderComposeRecipients(toList, ccList);
        _composeAttachments = attachments;
        renderComposeAttachments(_composeAttachments);
        document.getElementById("btn-refresh").classList.remove("spinning");
    });
}

function renderComposeRecipients(toList, ccList) {
    const area  = document.getElementById("compose-recipients");
    const badge = document.getElementById("recipients-count");
    const total = toList.length + ccList.length;
    badge.textContent = total || "";

    if (total === 0) {
        area.innerHTML = `<div class="compose-empty">No recipients yet. Add To / CC addresses then click &#8635; Refresh.</div>`;
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

    list.innerHTML = "";

    if (isComposeBulkMode()) {
        // Bundle mode: primary + supporting selection (same layout as read mode)
        document.getElementById("compose-bundle-footer").classList.remove("hidden");
        document.getElementById("btn-compose-upload").disabled = false;
        renderBundleList(attachments, list);
    } else {
        // Individual mode: each attachment gets its own "Upload & Share" button
        document.getElementById("compose-bundle-footer").classList.add("hidden");
        renderIndividualComposeList(attachments, list);
    }
}

function renderIndividualComposeList(attachments, container) {
    attachments.forEach((att, index) => {
        const div = document.createElement("div");
        div.className = "att-item";
        div.innerHTML = `
            <div class="att-individual-row">
                <div class="att-info">
                    <div class="att-name" title="${att.name}">${att.name}</div>
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

// Bundle upload in compose: upload primary + supporting, then share API + insert link
async function handleComposeBundleUpload() {
    const attachments  = _composeAttachments;
    const primaryRadio = document.querySelector("input[name='primaryIndex']:checked");
    if (!primaryRadio) { showComposeStatus("Please select a primary document."); return; }

    const primaryIndex     = parseInt(primaryRadio.value);
    const primaryAtt       = attachments[primaryIndex];
    const secondaryIndices = Array.from(
        document.querySelectorAll("input[name='secondaryIndex']:checked")
    ).map(c => parseInt(c.value)).filter(i => i !== primaryIndex);

    const uploadBtn = document.getElementById("btn-compose-upload");
    uploadBtn.disabled = true;
    document.getElementById("compose-result").classList.add("hidden");

    try {
        showComposeStatus("Signing in\u2026");
        const token = await getAuthToken();

        showComposeStatus("Uploading primary document\u2026");
        const blob = await getAttachmentBlob(primaryAtt.id, primaryAtt.name);
        const form = new FormData();
        form.append("document", blob, primaryAtt.name);

        const uploadResp = await fetch(UPLOAD_URL, {
            method: "POST", headers: { Authorization: "Bearer " + token }, body: form,
        });
        if (!uploadResp.ok)
            throw new Error("Upload failed (" + uploadResp.status + "): " + await uploadResp.text());

        const uploadData     = await uploadResp.json();
        const conversationId = uploadData.conversation_id || uploadData.conversationId || "";
        const docId          = uploadData.doc_id || uploadData.documentId || uploadData.id || "";

        if (!conversationId && !docId)
            throw new Error("No conversation ID or document ID returned by upload.");

        if (secondaryIndices.length > 0) {
            showComposeStatus("Uploading " + secondaryIndices.length + " supporting doc(s)\u2026");
            for (const idx of secondaryIndices) {
                await uploadSupportingById(_composeAttachments[idx], conversationId, token);
            }
        }

        showComposeStatus("Creating share link\u2026");
        const shareLink = await callShareApi(token, conversationId, docId);

        showComposeStatus("Inserting link into email\u2026");
        const documentURL = `${PROXY_BASE}/conversation?conversation-id=${conversationId}&doc-id=${docId}`;
        await insertShareLinkIntoBody(documentURL, primaryAtt.name);

        showComposeStatus("");
        renderComposeResult(shareLink);

    } catch (err) {
        console.error("Compose bundle upload error:", err);
        showComposeStatus("Error: " + err.message);
        _cachedSmartBlueToken = null;
    } finally {
        uploadBtn.disabled = false;
    }
}

// Individual upload in compose: one doc → share API → insert link
async function handleComposeSingleUpload(index) {
    const att = _composeAttachments[index];
    document.querySelectorAll(".btn-upload-share").forEach(b => b.disabled = true);
    document.getElementById("compose-result").classList.add("hidden");
    showComposeStatus("Signing in\u2026");

    try {
        const token = await getAuthToken();

        showComposeStatus("Uploading " + att.name + "\u2026");
        const blob = await getAttachmentBlob(att.id, att.name);
        const form = new FormData();
        form.append("document", blob, att.name);

        const uploadResp = await fetch(UPLOAD_URL, {
            method: "POST", headers: { Authorization: "Bearer " + token }, body: form,
        });
        if (!uploadResp.ok)
            throw new Error("Upload failed (" + uploadResp.status + "): " + await uploadResp.text());

        const uploadData     = await uploadResp.json();
        const conversationId = uploadData.conversation_id || uploadData.conversationId || "";
        const docId          = uploadData.doc_id || uploadData.documentId || uploadData.id || "";

        if (!conversationId && !docId)
            throw new Error("No document ID returned by upload.");

        showComposeStatus("Creating share link\u2026");
        const shareLink = await callShareApi(token, conversationId, docId);

        showComposeStatus("Inserting link into email\u2026");
        const documentURL = `${PROXY_BASE}/conversation?conversation-id=${conversationId}&doc-id=${docId}`;
        await insertShareLinkIntoBody(documentURL, att.name);

        showComposeStatus("");
        renderComposeResult(shareLink);

    } catch (err) {
        console.error("Compose single upload error:", err);
        showComposeStatus("Error: " + err.message);
        _cachedSmartBlueToken = null;
    } finally {
        document.querySelectorAll(".btn-upload-share").forEach(b => b.disabled = false);
    }
}

async function callShareApi(token, conversationId, docId) {
    const payload = { sender_email: _senderEmail, recipients: _composeRecipients };
    if (conversationId) payload.conversation_id = conversationId;
    if (docId)          payload.doc_id           = docId;

    const resp = await fetch(SHARE_URL, {
        method: "POST",
        headers: { "Content-Type": "application/json", Authorization: "Bearer " + token },
        body:   JSON.stringify(payload),
    });
    if (!resp.ok)
        throw new Error("Share API failed (" + resp.status + "): " + await resp.text());

    const data = await resp.json();
    const url  = data.share_url || data.shareUrl || data.url || "";
    if (!url) throw new Error("Share API returned no URL.");
    return url;
}

function insertShareLinkIntoBody(link, filename) {
    return new Promise((resolve) => {
        const html = `<p style="font-family:sans-serif;margin:8px 0;">`
                   + `<a href="${link}" target="_blank" `
                   + `style="color:#0D47A1;font-weight:600;text-decoration:none;">`
                   + `\uD83D\uDCC4 ${filename} \u2014 View on SmartBlue</a></p>`;

        Office.context.mailbox.item.body.setSelectedDataAsync(
            html,
            { coercionType: Office.CoercionType.Html },
            (result) => {
                if (result.status === Office.AsyncResultStatus.Failed) {
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

function showComposeStatus(msg) { document.getElementById("compose-status").innerText = msg; }

// ══════════════════════════════════════════════════════════════════════════
// SHARED ATTACHMENT HELPERS
// ══════════════════════════════════════════════════════════════════════════

// Shared bundle list renderer — used by both read and compose bundle modes
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
    if (e.target.name === "secondaryIndex")
        e.target.dataset.userUnchecked = e.target.checked ? "" : "1";
});

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
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body:   JSON.stringify({ idToken }),
    });
    if (!authResp.ok) throw new Error("Auth failed (" + authResp.status + "): " + await authResp.text());

    const { token } = await authResp.json();
    if (!token) throw new Error("No token returned from auth proxy.");
    _cachedSmartBlueToken = token;
    return token;
}

// ══════════════════════════════════════════════════════════════════════════
// UPLOAD HELPERS
// ══════════════════════════════════════════════════════════════════════════
async function uploadPrimary(att, token) {
    const blob = await getAttachmentBlob(att.id, att.name);
    const form = new FormData();
    form.append("document", blob, att.name);

    const resp = await fetch(UPLOAD_URL, {
        method: "POST", headers: { Authorization: "Bearer " + token }, body: form,
    });
    if (!resp.ok) throw new Error("Upload failed (" + resp.status + "): " + await resp.text());

    const data           = await resp.json();
    const conversationId = data.conversation_id || data.conversationId;
    const documentId     = data.doc_id || data.documentId || data.id || null;

    if (!conversationId) throw new Error("No conversation_id returned by upload.");
    return { conversationId, documentId };
}

async function uploadSupporting(att, conversationId, token) {
    return uploadSupportingById(att, conversationId, token);
}

async function uploadSupportingById(att, conversationId, token) {
    const blob = await getAttachmentBlob(att.id, att.name);
    const form = new FormData();
    form.append("document", blob, att.name);

    const resp = await fetch(
        `${BUNDLE_ADD_URL}?conversation_id=${encodeURIComponent(conversationId)}`,
        { method: "POST", headers: { Authorization: "Bearer " + token }, body: form }
    );
    if (!resp.ok) console.warn("Supporting upload failed:", att.name, await resp.text());
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
// CHAT
// ══════════════════════════════════════════════════════════════════════════
async function enterChat(conversationId, documentId, token) {
    currentConversationId = conversationId;
    currentDocumentId     = documentId;

    document.getElementById("view-read-init").classList.add("hidden");
    document.getElementById("view-read").classList.add("hidden");
    document.getElementById("view-compose").classList.add("hidden");
    document.getElementById("view-chat").classList.remove("hidden");
    document.getElementById("btn-back").classList.remove("hidden");
    document.getElementById("chat-history").innerHTML = "";
    hideSuggestions();

    showTypingIndicator();

    // documentId is optional when opening from a share link —
    // the conversation already exists so only conversationId is needed
    const welcomePayload = { conversationId };
    if (documentId) welcomePayload.documentId = documentId;

    try {
        const resp = await fetch(WELCOME_URL, {
            method:  "POST",
            headers: { "Content-Type": "application/json", Authorization: "Bearer " + token },
            body:    JSON.stringify(welcomePayload),
        });

        const rawText = await resp.text();
        hideTypingIndicator();

        if (!resp.ok) {
            console.error("Welcome API failed:", resp.status, rawText);
            appendMessage("ai", "Could not load welcome message (" + resp.status + "). You can still ask questions below.");
            return;
        }

        let data;
        try { data = JSON.parse(rawText); } catch {
            console.error("Welcome response not JSON:", rawText);
            appendMessage("ai", "Hello! How can I help you with this document?");
            return;
        }

        console.log("Welcome response:", data);

        // SmartBlue may return the message under various field names
        const welcomeMsg =
            data.answer       ||
            data.response     ||
            data.message      ||
            data.text         ||
            data.content      ||
            data.welcomeText  ||
            data.welcome_text ||
            (typeof data === "string" ? data : null);

        if (welcomeMsg) {
            appendMessage("ai", welcomeMsg);
        } else {
            console.warn("Welcome API returned no recognised text field. Keys:", Object.keys(data));
            appendMessage("ai", "Hello! How can I help you with this document?");
        }

        const tags = Array.isArray(data.tags) ? data.tags : [];
        if (tags.length) renderSuggestions(tags);

    } catch (err) {
        hideTypingIndicator();
        console.error("Welcome fetch error:", err);
        appendMessage("ai", "Network error loading welcome message. You can still ask questions below.");
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
            method: "POST",
            headers: { "Content-Type": "application/json", Authorization: "Bearer " + token },
            body:   JSON.stringify({ conversationId: currentConversationId, text, isMobile: false }),
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