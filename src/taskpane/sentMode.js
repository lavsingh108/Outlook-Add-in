// ── Sent Mode ────────────────────────────────────────────────────────────────
// Handles the taskpane when the user is viewing a sent email.
// Scans the email body for BlueAI shared-document links and, when found,
// renders an "Add to Bundle" card so the user can add the shared document
// to any existing conversation.

import { escHtml, parseDocUrl } from "./utils/helpers.js";
import { getAuthToken, clearToken } from "./services/authService.js";
import {
    loadCustomProps, getConversationMap,
    saveConversationRecord, getThreadContextAll, saveThreadContext,
} from "./services/storageService.js";
import { uploadSupportingById } from "./services/uploadService.js";
import { resolveShareId } from "./services/apiService.js";
import { showReadStatus } from "./ui/chatUI.js";
import { setCustomProps, setChatFromSent } from "./state.js";
import { _customProps } from "./state.js";
import { enterChat } from "./chat.js";

// ── Share-link extraction (same logic as readMode) ─────────────────────────

function extractShareLinksFromBody() {
    return new Promise((resolve, reject) => {
        Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, result => {
            if (result.status !== Office.AsyncResultStatus.Succeeded) {
                reject(new Error(result.error?.message || "Body read failed"));
                return;
            }
            const html    = result.value || "";
            const found   = [];
            const seen    = new Set();

            const anchorRe = /<a[^>]+href=["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>/gi;
            let m;
            while ((m = anchorRe.exec(html)) !== null) {
                const linkText = m[2].replace(/<[^>]+>/g, "").replace(/\s+/g, " ").trim();
                const parsed   = parseDocUrl(m[1]);
                const key      = parsed.shareId || parsed.conversationId;
                if (key && !seen.has(key)) {
                    seen.add(key);
                    found.push({ ...parsed, linkText: linkText || m[1] });
                }
            }

            // Also scan bare URLs (non-anchor)
            const urlRe = /https?:\/\/[^\s"'<>)]+/gi;
            while ((m = urlRe.exec(html)) !== null) {
                const parsed = parseDocUrl(m[0]);
                const key    = parsed.shareId || parsed.conversationId;
                if (key && !seen.has(key)) {
                    seen.add(key);
                    found.push({ ...parsed, linkText: null });
                }
            }

            resolve(found);
        });
    });
}

// ── Conversation picker helpers ────────────────────────────────────────────

function getAllConversations(customProps) {
    const map = {};
    if (customProps) {
        Object.values(getConversationMap(customProps))
            .filter(r => r.uploadType !== "shared-link")
            .forEach(r => { map[r.conversationId] = r; });
    }
    getThreadContextAll()
        .filter(r => r.uploadType !== "shared-link")
        .forEach(r => { if (!map[r.conversationId]) map[r.conversationId] = r; });
    return Object.values(map).sort((a, b) => (b.timestamp || 0) - (a.timestamp || 0));
}

// ── Render a single shared-document card ──────────────────────────────────

function renderSentShareCard(shareInfo, index, listEl, customProps) {
    const displayText = shareInfo.linkText || shareInfo.shareUrl || "BlueAI Document";
    const displayUrl  = shareInfo.shareUrl || "";

    const card = document.createElement("div");
    card.className = "sent-share-card";
    card.dataset.index = index;

    card.innerHTML = `
        <div class="read-share-inner">
            <svg class="read-share-file-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor"
                 stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                <polyline points="14 2 14 8 20 8"/>
            </svg>
            <div class="read-share-info">
                <div class="read-share-name" title="${escHtml(displayText)}">${escHtml(displayText)}</div>
                <div class="read-share-url"  title="${escHtml(displayUrl)}">${escHtml(displayUrl)}</div>
            </div>
        </div>
        <div class="sent-card-actions">
            <button class="btn-add-to-bundle" data-index="${index}">
                <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor"
                     stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
                    <line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/>
                </svg>
                Add to Bundle
            </button>
        </div>`;

    card.querySelector(".btn-add-to-bundle").onclick = () =>
        handleAddToBundle(shareInfo, card, customProps);

    listEl.appendChild(card);
}

// ── "Add to Bundle" click handler ─────────────────────────────────────────

async function handleAddToBundle(shareInfo, cardEl, customProps) {
    const btn = cardEl.querySelector(".btn-add-to-bundle");
    btn.disabled = true;
    showSentStatus("Signing in\u2026");

    try {
        const token = await getAuthToken();

        // Resolve share-id → conversationId + docId if needed
        let { conversationId, docId } = shareInfo;
        if (!conversationId && shareInfo.shareId) {
            showSentStatus("Resolving share link\u2026");
            const resolved = await resolveShareId(shareInfo.shareId, token);
            conversationId = resolved.conversationId;
            docId          = resolved.docId;
        }

        // Pick the conversation to add to
        const conversations = getAllConversations(customProps);

        if (conversations.length === 0) {
            // No existing conversations — upload as a new primary document
            showSentStatus("No existing conversation found. Uploading as new document\u2026");
            await _uploadShareAsNewDoc(shareInfo, conversationId, docId, token, cardEl);
            return;
        }

        if (conversations.length === 1) {
            // Only one conversation — add directly
            await _doAddToBundle(shareInfo, conversations[0], conversationId, docId, token, cardEl);
        } else {
            // Multiple conversations — show a picker
            showConversationPicker(cardEl, conversations, shareInfo, conversationId, docId, token);
        }
    } catch (err) {
        console.error("Add to bundle (sent) error:", err);
        showSentStatus("Error: " + err.message);
        clearToken();
        btn.disabled = false;
    }
}

// Show an inline conversation picker when there are multiple options
function showConversationPicker(cardEl, conversations, shareInfo, conversationId, docId, token) {
    // Remove any existing picker
    const existingPicker = cardEl.querySelector(".sent-conv-picker");
    if (existingPicker) existingPicker.remove();

    const picker = document.createElement("div");
    picker.className = "sent-conv-picker";
    picker.innerHTML = `<div class="sent-picker-label">Add to which conversation?</div>`;

    conversations.forEach(conv => {
        const date = conv.timestamp
            ? new Date(conv.timestamp).toLocaleDateString(undefined, { month: "short", day: "numeric" })
            : "";
        const btn = document.createElement("button");
        btn.className = "sent-picker-conv-btn";
        btn.innerHTML = `
            <span class="sent-picker-conv-name">${escHtml(conv.label || "Document")}</span>
            <span class="sent-picker-conv-meta">${escHtml(conv.uploadType || "")}${date ? " \u00b7 " + date : ""}</span>`;
        btn.onclick = async () => {
            picker.remove();
            try {
                await _doAddToBundle(shareInfo, conv, conversationId, docId, token, cardEl);
            } catch (err) {
                showSentStatus("Error: " + err.message);
                clearToken();
                const addBtn = cardEl.querySelector(".btn-add-to-bundle");
                if (addBtn) addBtn.disabled = false;
            }
        };
        picker.appendChild(btn);
    });

    cardEl.appendChild(picker);
    showSentStatus("");
}

// Fetch the shared doc as a Blob via its URL and add it to an existing conversation
async function _doAddToBundle(shareInfo, targetConv, conversationId, docId, token, cardEl) {
    const { conversationId: existingConvId, documentId: existingDocId } = targetConv;
    const label = shareInfo.linkText || shareInfo.shareUrl || "Shared Document";

    showSentStatus("Fetching shared document\u2026");

    // Download the shared document content
    const shareUrl = shareInfo.shareUrl;
    if (!shareUrl) throw new Error("No share URL available to download.");

    const docBlob = await _fetchSharedDocBlob(shareUrl, token);

    showSentStatus("Adding to bundle\u2026");

    const { BUNDLE_ADD_URL } = await import("./config.js");
    const form = new FormData();
    form.append("document", docBlob, label.endsWith(".pdf") ? label : label + ".pdf");

    const resp = await fetch(
        `${BUNDLE_ADD_URL}?conversation_id=${encodeURIComponent(existingConvId)}`,
        { method: "POST", headers: { Authorization: "Bearer " + token }, body: form }
    );
    if (!resp.ok) throw new Error("Bundle add failed (" + resp.status + "): " + await resp.text());

    // Save record
    const record = {
        conversationId: existingConvId,
        documentId:     existingDocId,
        label,
        uploadType:     "bundle",
        timestamp:      Date.now(),
    };
    saveThreadContext(record);

    showSentStatus("");
    _markCardAdded(cardEl, label, existingConvId, existingDocId, token);
}

async function _uploadShareAsNewDoc(shareInfo, conversationId, docId, token, cardEl) {
    const { UPLOAD_URL } = await import("./config.js");
    const label    = shareInfo.linkText || shareInfo.shareUrl || "Shared Document";
    const docBlob  = await _fetchSharedDocBlob(shareInfo.shareUrl, token);
    const form     = new FormData();
    form.append("document", docBlob, label);

    const resp = await fetch(UPLOAD_URL, {
        method:  "POST",
        headers: { Authorization: "Bearer " + token },
        body:    form,
    });
    if (!resp.ok) throw new Error("Upload failed (" + resp.status + "): " + await resp.text());

    const data = await resp.json();
    const newConvId = data.conversation_id || data.conversationId;
    const newDocId  = data.doc_id || data.document_id || data.documentId || data.id;

    const record = { conversationId: newConvId, documentId: newDocId, label, uploadType: "single", timestamp: Date.now() };
    saveThreadContext(record);

    showSentStatus("");
    _markCardAdded(cardEl, label, newConvId, newDocId, token);
}

// Attempt to download a shared BlueAI doc blob (may be protected — best effort)
async function _fetchSharedDocBlob(shareUrl, token) {
    const resp = await fetch(shareUrl, {
        headers: { Authorization: "Bearer " + token },
    });
    if (!resp.ok) throw new Error("Could not fetch shared document (" + resp.status + "). The document may require separate access.");
    return await resp.blob();
}

// Replace the action area with a success state + "Open Chat" button
function _markCardAdded(cardEl, label, convId, docId, token) {
    const actionsEl = cardEl.querySelector(".sent-card-actions");
    actionsEl.innerHTML = `
        <span class="sent-added-badge">
            <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor"
                 stroke-width="3" stroke-linecap="round" stroke-linejoin="round">
                <polyline points="20 6 9 17 4 12"/>
            </svg>
            Added
        </span>
        <button class="btn-open-chat-sent">
            <svg width="11" height="11" viewBox="0 0 24 24" fill="currentColor">
                <polygon points="5 3 19 12 5 21 5 3"/>
            </svg>
            Open Chat
        </button>`;

    actionsEl.querySelector(".btn-open-chat-sent").onclick = async () => {
        const btn = actionsEl.querySelector(".btn-open-chat-sent");
        btn.disabled = true;
        showSentStatus("Opening chat\u2026");
        try {
            setChatFromSent(true);
            document.getElementById("view-sent").classList.add("hidden");
            const t = token || await getAuthToken();
            await enterChat(convId, docId, t);
        } catch (err) {
            showSentStatus("Error: " + err.message);
            clearToken();
            setChatFromSent(false);
            document.getElementById("view-sent").classList.remove("hidden");
            btn.disabled = false;
        }
    };
}

// ── Status helper ──────────────────────────────────────────────────────────

function showSentStatus(msg) {
    const el = document.getElementById("sent-status-msg");
    if (el) el.innerText = msg;
}

// ── Init ───────────────────────────────────────────────────────────────────

export function initSent() {
    document.querySelector(".header-title").textContent = "Sent Mail";

    // Wire back button
    document.getElementById("btn-back").onclick = () =>
        import("./navigation.js").then(m => m.switchToReadView());

    // Wire send chat button
    document.getElementById("btn-send").onclick = () =>
        import("./chat.js").then(m => m.sendChatMessage());
    document.getElementById("user-input").addEventListener("keydown", e => {
        if (e.key === "Enter" && !e.shiftKey) {
            e.preventDefault();
            import("./chat.js").then(m => m.sendChatMessage());
        }
    });

    // Show loading screen
    const initView   = document.getElementById("view-sent-init");
    const sentView   = document.getElementById("view-sent");
    initView.classList.remove("hidden");

    extractShareLinksFromBody()
        .then(async shareLinks => {
            initView.classList.add("hidden");
            sentView.classList.remove("hidden");

            const listEl    = document.getElementById("sent-share-list");
            const noDocEl   = document.getElementById("sent-no-doc");
            const sectionEl = document.getElementById("sent-share-section");

            const blueLinks = shareLinks.filter(s => s.conversationId || s.shareId);

            if (!blueLinks.length) {
                noDocEl.classList.remove("hidden");
                sectionEl.classList.add("hidden");
                return;
            }

            noDocEl.classList.add("hidden");
            sectionEl.classList.remove("hidden");

            // Load custom props for conversation lookup
            let cp = null;
            try { cp = await loadCustomProps(); setCustomProps(cp); } catch (_) {}

            blueLinks.forEach((shareInfo, i) =>
                renderSentShareCard(shareInfo, i, listEl, cp)
            );
        })
        .catch(err => {
            initView.classList.add("hidden");
            sentView.classList.remove("hidden");
            document.getElementById("sent-no-doc").classList.remove("hidden");
            document.getElementById("sent-no-doc").textContent = "Error reading email: " + err.message;
        });
}
