import { BLUE_BASE } from "./config.js";
import { state } from "./state.js";
import { getAuthToken, clearToken } from "./services/authService.js";
import { uploadPrimary, uploadSupportingById } from "./services/uploadService.js";
import { callShareApi } from "./services/shareService.js";
import { fetchHistory, fetchWelcome, askQuestion } from "./services/chatService.js";
import { extractShareLinkFromBody } from "./utils/urlUtils.js";
import { escHtml } from "./utils/domUtils.js";
import { fallbackCopy } from "./utils/clipboardUtils.js";
import {
    showTypingIndicator, hideTypingIndicator, appendMessage,
    renderSuggestions, hideSuggestions, restoreConversationHistory,
} from "./helpers/chatHelpers.js";
import {
    renderBundleList,
    renderIndividualReadList,
    renderIndividualComposeList,
} from "./helpers/attachmentHelpers.js";
import { showReadStatus, showReadInitError, showComposeStatus } from "./helpers/statusHelpers.js";
import { renderShareSection, insertShareLinkIntoBody, renderComposeResult } from "./helpers/composeHelpers.js";
import {
    loadCustomProps, getConversationMap,
    saveConversationRecord, singleFingerprint, bundleFingerprint,
} from "./utils/customPropsUtils.js";

// Compose-specific state (local to this module)
let _composeAttachments = [];
let _composeRecipients  = [];
let _senderEmail        = "";

// Custom Properties handle — loaded once in initRead, reused across calls
let _customProps = null;

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
                renderShareSection(shareInfo, async (info) => {
                    showReadStatus("Signing in…");
                    try {
                        const token = await getAuthToken();
                        await enterChat(info.conversationId, info.docId, token);
                    } catch (err) {
                        showReadStatus("Error: " + err.message);
                        clearToken();
                        throw err;
                    }
                });
            }

            loadCustomProps()
                .then(cp => {
                    _customProps = cp;
                    renderPreviousChats();
                })
                .catch(() => { /* non-fatal — previous chats section stays hidden */ });

            loadReadAttachments();

            const atts = Office.context.mailbox.item.attachments || [];
            if (shareInfo && shareInfo.conversationId && atts.length > 0) {
                document.getElementById("read-or-divider").classList.remove("hidden");
            }
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

    const map     = getConversationMap(_customProps);
    const records = Object.values(map);
    if (!records.length) { section.classList.add("hidden"); return; }

    list.innerHTML = "";
    records
        .slice()
        .sort((a, b) => (b.timestamp || 0) - (a.timestamp || 0))
        .forEach(rec => {
            const item = document.createElement("div");
            item.className = "prev-chat-item";
            const date = rec.timestamp
                ? new Date(rec.timestamp).toLocaleDateString(undefined, { month: "short", day: "numeric" })
                : "";
            item.innerHTML = `
                <div class="prev-chat-info">
                    <div class="prev-chat-name" title="${escHtml(rec.label || "")}">${escHtml(rec.label || "Document")}</div>
                    <div class="prev-chat-meta">${escHtml(rec.uploadType || "")}${date ? " · " + date : ""}</div>
                </div>
                <button class="btn-resume">Resume</button>`;
            item.querySelector(".btn-resume").onclick = async () => {
                showReadStatus("Signing in…");
                try {
                    const token = await getAuthToken();
                    await enterChat(rec.conversationId, rec.documentId, token);
                } catch (err) {
                    showReadStatus("Error: " + err.message);
                    clearToken();
                }
            };
            list.appendChild(item);
        });

    section.classList.remove("hidden");
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
        renderIndividualReadList(attachments, listDiv, handleReadSingleUpload);
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
    const secondaryAtts    = secondaryIndices.map(i => attachments[i]);

    document.getElementById("btn-upload-bundle").disabled = true;
    showReadStatus("Signing in…");

    try {
        const token = await getAuthToken();

        // Resume from cache if already uploaded
        if (_customProps) {
            const fp  = bundleFingerprint(primaryAtt, secondaryAtts);
            const rec = getConversationMap(_customProps)[fp];
            if (rec) {
                showReadStatus("");
                return await enterChat(rec.conversationId, rec.documentId, token);
            }
        }

        showReadStatus("Uploading primary document…");
        const { conversationId, documentId } = await uploadPrimary(primaryAtt, token);

        if (secondaryAtts.length > 0) {
            showReadStatus("Uploading " + secondaryAtts.length + " supporting doc(s)…");
            for (const att of secondaryAtts) {
                await uploadSupportingById(att, conversationId, token);
            }
        }

        if (_customProps) {
            const fp = bundleFingerprint(primaryAtt, secondaryAtts);
            saveConversationRecord(_customProps, fp, {
                conversationId, documentId,
                label: primaryAtt.name,
                uploadType: "bundle",
                timestamp: Date.now(),
            }).catch(err => console.warn("customProps save failed:", err.message));
        }

        await enterChat(conversationId, documentId, token);
    } catch (err) {
        console.error("Read bundle upload error:", err);
        showReadStatus("Error: " + err.message);
        clearToken();
        document.getElementById("btn-upload-bundle").disabled = false;
    }
}

async function handleReadSingleUpload(index) {
    const att = Office.context.mailbox.item.attachments[index];
    document.querySelectorAll(".btn-upload-single").forEach(b => b.disabled = true);
    showReadStatus("Signing in…");

    try {
        const token = await getAuthToken();

        // Resume from cache if already uploaded
        if (_customProps) {
            const fp  = singleFingerprint(att);
            const rec = getConversationMap(_customProps)[fp];
            if (rec) {
                showReadStatus("");
                return await enterChat(rec.conversationId, rec.documentId, token);
            }
        }

        showReadStatus("Uploading " + att.name + "…");
        const { conversationId, documentId } = await uploadPrimary(att, token);

        if (_customProps) {
            saveConversationRecord(_customProps, singleFingerprint(att), {
                conversationId, documentId,
                label: att.name,
                uploadType: "single",
                timestamp: Date.now(),
            }).catch(err => console.warn("customProps save failed:", err.message));
        }

        await enterChat(conversationId, documentId, token);
    } catch (err) {
        console.error("Read single upload error:", err);
        showReadStatus("Error: " + err.message);
        clearToken();
        document.querySelectorAll(".btn-upload-single").forEach(b => b.disabled = false);
    }
}

function switchToReadView() {
    document.getElementById("view-chat").classList.add("hidden");
    document.getElementById("view-read").classList.remove("hidden");
    document.getElementById("btn-back").classList.add("hidden");
    document.getElementById("chat-history").innerHTML = "";
    hideSuggestions();
    state.currentConversationId = null;
    state.currentDocumentId     = null;
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
        const row  = document.createElement("div");
        row.className = "recipient-row";
        const lbl  = document.createElement("span");
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
        document.getElementById("compose-bundle-footer").classList.remove("hidden");
        document.getElementById("btn-compose-upload").disabled = false;
        renderBundleList(attachments, list);
    } else {
        document.getElementById("compose-bundle-footer").classList.add("hidden");
        renderIndividualComposeList(attachments, list, handleComposeSingleUpload);
    }
}

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
        showComposeStatus("Signing in…");
        const token = await getAuthToken();

        showComposeStatus("Uploading primary document…");
        const { conversationId, documentId } = await uploadPrimary(primaryAtt, token);

        if (secondaryIndices.length > 0) {
            showComposeStatus("Uploading " + secondaryIndices.length + " supporting doc(s)…");
            for (const idx of secondaryIndices) {
                await uploadSupportingById(_composeAttachments[idx], conversationId, token);
            }
        }

        showComposeStatus("Creating share link…");
        await callShareApi(token, conversationId, documentId, _senderEmail, _composeRecipients);

        showComposeStatus("Inserting link into email…");
        const documentURL = `${BLUE_BASE}/conversation?conversation-id=${conversationId}&doc-id=${documentId}`;
        await insertShareLinkIntoBody(documentURL, primaryAtt.name);

        showComposeStatus("");
        renderComposeResult(documentURL);

    } catch (err) {
        console.error("Compose bundle upload error:", err);
        showComposeStatus("Error: " + err.message);
        clearToken();
    } finally {
        uploadBtn.disabled = false;
    }
}

async function handleComposeSingleUpload(index) {
    const att = _composeAttachments[index];
    document.querySelectorAll(".btn-upload-share").forEach(b => b.disabled = true);
    document.getElementById("compose-result").classList.add("hidden");
    showComposeStatus("Signing in…");

    try {
        const token = await getAuthToken();

        showComposeStatus("Uploading " + att.name + "…");
        const { conversationId, documentId } = await uploadPrimary(att, token);

        showComposeStatus("Creating share link…");
        await callShareApi(token, conversationId, documentId, _senderEmail, _composeRecipients);

        showComposeStatus("Inserting link into email…");
        const documentURL = `${BLUE_BASE}/conversation?conversation-id=${conversationId}&doc-id=${documentId}`;
        await insertShareLinkIntoBody(documentURL, att.name);

        showComposeStatus("");
        renderComposeResult(documentURL);

    } catch (err) {
        console.error("Compose single upload error:", err);
        showComposeStatus("Error: " + err.message);
        clearToken();
    } finally {
        document.querySelectorAll(".btn-upload-share").forEach(b => b.disabled = false);
    }
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
        appendMessage("ai", "Cannot start chat: document ID is missing. Please re-share the document.");
        return;
    }

    showTypingIndicator();

    // Step 1: try to restore existing conversation history
    let historyRestored = false;
    try {
        const histResp = await fetchHistory(token, conversationId, documentId);
        if (histResp.ok) {
            const conv  = await histResp.json();
            const msgs  = Array.isArray(conv.messages) ? conv.messages : [];
            const hasAI = msgs.some(m => m.sender === "assistant" || m.role === "assistant");
            if (hasAI) {
                hideTypingIndicator();
                restoreConversationHistory(msgs, sendChatMessage);
                historyRestored = true;
            }
        } else {
            console.warn("History API returned", histResp.status, "— falling through to welcome");
        }
    } catch (histErr) {
        console.warn("History fetch failed (non-fatal):", histErr.message, "— falling through to welcome");
    }

    if (historyRestored) return;

    // Step 2: no prior history — call welcome API
    try {
        const resp    = await fetchWelcome(token, conversationId, documentId);
        const rawText = await resp.text();
        hideTypingIndicator();

        if (!resp.ok) {
            console.error("Welcome API failed:", resp.status, rawText);
            appendMessage("ai", "Could not load welcome message (" + resp.status + "). You can still ask questions below.");
            return;
        }

        let data;
        try { data = JSON.parse(rawText); } catch {
            appendMessage("ai", "Hello! How can I help you with this document?");
            return;
        }

        const welcomeMsg =
            data.answer || data.response || data.message ||
            data.text   || data.content  || data.welcomeText ||
            data.welcome_text || (typeof data === "string" ? data : null);

        appendMessage("ai", welcomeMsg || "Hello! How can I help you with this document?");

        const tags = Array.isArray(data.tags) ? data.tags : [];
        if (tags.length) renderSuggestions(tags, sendChatMessage);

    } catch (err) {
        hideTypingIndicator();
        console.error("enterChat welcome error:", err);
        appendMessage("ai", "Network error. You can still ask questions below.");
    }
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
        const resp  = await askQuestion(token, state.currentConversationId, text);
        hideTypingIndicator();
        if (!resp.ok) throw new Error("Ask failed (" + resp.status + "): " + await resp.text());
        const data = await resp.json();
        appendMessage("ai", data.answer || data.response || "No response received.");
        const tags = Array.isArray(data.tags) ? data.tags : [];
        if (tags.length) renderSuggestions(tags, sendChatMessage);
    } catch (err) {
        hideTypingIndicator();
        appendMessage("ai", "Error: " + err.message);
        clearToken();
    } finally {
        document.getElementById("btn-send").disabled = false;
    }
}

// Track user manually unchecking secondary boxes across both modes
document.addEventListener("change", (e) => {
    if (e.target.name === "secondaryIndex")
        e.target.dataset.userUnchecked = e.target.checked ? "" : "1";
});
