// ── Read Mode ───────────────────────────────────────────────────────────────
// Handles the read (View Document) pane: attachment list, upload handlers,
// previous-chat list, and share-link rendering.

import { escHtml, parseDocUrl } from "./utils/helpers.js";
import { getAuthToken, clearToken } from "./services/authService.js";
import {
    loadCustomProps, getConversationMap,
    saveConversationRecord, getThreadContextAll, saveThreadContext,
    singleFingerprint, bundleFingerprint, getAttachmentRecord, isAttachmentUploaded,
} from "./services/storageService.js";
import { uploadPrimary, uploadSupportingById } from "./services/uploadService.js";
import { resolveShareId } from "./services/apiService.js";
import {
    renderBundleList, renderIndividualReadList,
    renderAddToBundleList, updateBundleSelection,
} from "./ui/attachmentUI.js";
import {
    showReadStatus, showReadInitError,
} from "./ui/chatUI.js";
import {
    state, _customProps, _readShareInfo,
    setCustomProps, setReadShareInfo,
} from "./state.js";
import { enterChat } from "./chat.js";

// ── Share link extraction from email body ──────────────────────────────────

function extractShareLinkFromBody() {
    return new Promise((resolve, reject) => {
        Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, result => {
            if (result.status !== Office.AsyncResultStatus.Succeeded) {
                reject(new Error(result.error?.message || "Body read failed"));
                return;
            }
            const html     = result.value || "";
            const anchorRe = /<a[^>]+href=["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>/gi;
            let m;
            while ((m = anchorRe.exec(html)) !== null) {
                const linkText = m[2].replace(/<[^>]+>/g, "").replace(/\s+/g, " ").trim();
                const parsed   = parseDocUrl(m[1]);
                if (parsed.conversationId || parsed.shareId) {
                    resolve({ ...parsed, linkText: linkText || m[1] });
                    return;
                }
            }
            const urlRe = /https?:\/\/[^\s"'<>)]+/gi;
            while ((m = urlRe.exec(html)) !== null) {
                const parsed = parseDocUrl(m[0]);
                if (parsed.conversationId || parsed.shareId) {
                    resolve({ ...parsed, linkText: null });
                    return;
                }
            }
            resolve({ conversationId: null, docId: null, shareId: null, linkText: null });
        });
    });
}

// ── Share section UI ───────────────────────────────────────────────────────

function renderShareSection(shareInfo) {
    const section     = document.getElementById("read-share-section");
    const card        = document.getElementById("read-share-card");
    const displayText = shareInfo.linkText || shareInfo.shareUrl || "View on SmartBlue";
    const displayUrl  = shareInfo.shareUrl  || "";

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
            let { conversationId, docId } = shareInfo;
            if (!conversationId && shareInfo.shareId) {
                showReadStatus("Resolving share link\u2026");
                const resolved  = await resolveShareId(shareInfo.shareId, token);
                conversationId  = resolved.conversationId;
                docId           = resolved.docId;
                shareInfo.conversationId = conversationId;
                shareInfo.docId          = docId;
            }
            await enterChat(conversationId, docId, token);
        } catch (err) {
            showReadStatus("Error: " + err.message);
            clearToken();
            btn.disabled = false;
        }
    };
}

// ── Previous chats ─────────────────────────────────────────────────────────

export function renderPreviousChats() {
    const section = document.getElementById("read-prev-section");
    const list    = document.getElementById("read-prev-list");
    if (!_customProps) { section.classList.add("hidden"); return; }

    // Merge custom-props records with thread roamingSettings records
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
    records.slice()
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
                    <div class="prev-chat-meta">${escHtml(rec.uploadType || "")}${date ? " \u00b7 " + date : ""}</div>
                </div>
                <button class="btn-resume">Resume</button>`;
            item.querySelector(".btn-resume").onclick = async () => {
                showReadStatus("Signing in\u2026");
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

// ── Attachment list rendering ──────────────────────────────────────────────

function isReadBulkMode() {
    return document.getElementById("chk-bulk").checked;
}

export function loadReadAttachments() {
    const attachments   = Office.context.mailbox.item.attachments || [];
    const listDiv       = document.getElementById("attachment-list");
    const footerDiv     = document.getElementById("bundle-footer");
    const attachSection = document.getElementById("read-attach-section");
    const divider       = document.getElementById("read-or-divider");

    if (!attachments.length) {
        attachSection?.classList.add("hidden");
        divider?.classList.add("hidden");
        return;
    }
    attachSection?.classList.remove("hidden");

    // hasContext: any uploaded record that is NOT a share link
    const cpRecords     = _customProps
        ? Object.values(getConversationMap(_customProps)).filter(r => r.uploadType !== "shared-link")
        : [];
    const threadRecords = getThreadContextAll().filter(r => r.uploadType !== "shared-link");
    const hasContext    = cpRecords.length > 0 || threadRecords.length > 0;

    if (isReadBulkMode()) {
        footerDiv.classList.remove("hidden");
        const bundleBtn = document.getElementById("btn-upload-bundle");
        bundleBtn.disabled = false;

        if (hasContext) {
            renderAddToBundleList(attachments, listDiv);
            const allAdded = attachments.every(a => isAttachmentUploaded(a, _customProps));
            if (allAdded) {
                bundleBtn.textContent = "\u2713 All Added";
                bundleBtn.disabled    = true;
                bundleBtn.classList.remove("btn-start-chat-att");
                bundleBtn.onclick     = null;
            } else {
                bundleBtn.textContent = "\uFF0B Add to Bundle";
                bundleBtn.disabled    = false;
                bundleBtn.classList.remove("btn-start-chat-att");
                bundleBtn.onclick     = handleReadAddToBundle;
            }
        } else {
            renderBundleList(attachments, listDiv, enterChat, handleReadBundleUpload);
            const primaryRec = getAttachmentRecord(attachments[0], _customProps);
            if (primaryRec) {
                bundleBtn.textContent = "\u25B6 Start Chat";
                bundleBtn.classList.add("btn-start-chat-att");
                bundleBtn.onclick = async () => {
                    bundleBtn.disabled = true;
                    showReadStatus("Signing in\u2026");
                    try {
                        const token = await getAuthToken();
                        await enterChat(primaryRec.conversationId, primaryRec.documentId, token);
                    } catch (err) {
                        showReadStatus("Error: " + err.message);
                        clearToken();
                        bundleBtn.disabled = false;
                    }
                };
            } else {
                bundleBtn.textContent = "\u2B06 Upload & Analyse";
                bundleBtn.classList.remove("btn-start-chat-att");
                bundleBtn.onclick = handleReadBundleUpload;
            }
        }
    } else {
        footerDiv.classList.add("hidden");
        renderIndividualReadList(
            attachments, listDiv, hasContext,
            enterChat, handleReadSingleUpload, handleReadAddToExisting
        );
    }
}

// ── Upload handlers ────────────────────────────────────────────────────────

async function handleReadBundleUpload() {
    const attachments  = Office.context.mailbox.item.attachments;
    const primaryRadio = document.querySelector("input[name='primaryIndex']:checked");
    if (!primaryRadio) { showReadStatus("Please select a primary document."); return; }

    const primaryIndex  = parseInt(primaryRadio.value);
    const primaryAtt    = attachments[primaryIndex];
    const secondaryAtts = Array.from(document.querySelectorAll("input[name='secondaryIndex']:checked"))
        .map(c => parseInt(c.value))
        .filter(i => i !== primaryIndex)
        .map(i => attachments[i]);

    document.getElementById("btn-upload-bundle").disabled = true;
    showReadStatus("Signing in\u2026");

    try {
        const token = await getAuthToken();
        if (_customProps) {
            const fp  = bundleFingerprint(primaryAtt, secondaryAtts);
            const rec = getConversationMap(_customProps)[fp];
            if (rec) { showReadStatus(""); return await enterChat(rec.conversationId, rec.documentId, token); }
        }

        showReadStatus("Uploading primary document\u2026");
        const { conversationId, documentId } = await uploadPrimary(primaryAtt, token);

        if (secondaryAtts.length > 0) {
            showReadStatus("Uploading " + secondaryAtts.length + " supporting doc(s)\u2026");
            for (const att of secondaryAtts) await uploadSupportingById(att, conversationId, token);
        }

        const record = { conversationId, documentId, label: primaryAtt.name, uploadType: "bundle", timestamp: Date.now() };
        if (_customProps) {
            saveConversationRecord(_customProps, bundleFingerprint(primaryAtt, secondaryAtts), record)
                .catch(err => console.warn("customProps save failed:", err.message));
        }
        saveThreadContext(record);
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
    document.querySelectorAll(".btn-upload-single").forEach(b => { b.disabled = true; });
    showReadStatus("Signing in\u2026");

    try {
        const token = await getAuthToken();
        if (_customProps) {
            const rec = getConversationMap(_customProps)[singleFingerprint(att)];
            if (rec) { showReadStatus(""); return await enterChat(rec.conversationId, rec.documentId, token); }
        }

        showReadStatus("Uploading " + att.name + "\u2026");
        const { conversationId, documentId } = await uploadPrimary(att, token);

        const record = { conversationId, documentId, label: att.name, uploadType: "single", timestamp: Date.now() };
        if (_customProps) {
            saveConversationRecord(_customProps, singleFingerprint(att), record)
                .catch(err => console.warn("customProps save failed:", err.message));
        }
        saveThreadContext(record);
        await enterChat(conversationId, documentId, token);
    } catch (err) {
        console.error("Read single upload error:", err);
        showReadStatus("Error: " + err.message);
        clearToken();
        document.querySelectorAll(".btn-upload-single").forEach(b => { b.disabled = false; });
    }
}

async function handleReadAddToBundle() {
    const attachments   = Office.context.mailbox.item.attachments;
    const secondaryAtts = Array.from(document.querySelectorAll("input[name='addToBundleIndex']:checked"))
        .map(c => parseInt(c.value))
        .map(i => attachments[i]);
    if (!secondaryAtts.length) { showReadStatus("Select at least one document to add."); return; }

    const latestRecord = _getLatestNonShareRecord();
    if (!latestRecord) { showReadStatus("No existing conversation found."); return; }
    const { conversationId: existingConvId, documentId: existingDocId } = latestRecord;

    document.getElementById("btn-upload-bundle").disabled = true;
    showReadStatus("Signing in\u2026");

    try {
        const token = await getAuthToken();
        showReadStatus("Adding " + secondaryAtts.length + " doc(s) to conversation\u2026");
        for (const att of secondaryAtts) await uploadSupportingById(att, existingConvId, token);

        if (_customProps) {
            await Promise.all(secondaryAtts.map(att =>
                saveConversationRecord(_customProps, singleFingerprint(att), {
                    conversationId: existingConvId, documentId: existingDocId,
                    label: att.name, uploadType: "bundle", timestamp: Date.now(),
                }).catch(err => console.warn("customProps save failed:", err.message))
            ));
        }

        showReadStatus("");
        await enterChat(existingConvId, existingDocId, token);
    } catch (err) {
        console.error("Add to bundle error:", err);
        showReadStatus("Error: " + err.message);
        clearToken();
        document.getElementById("btn-upload-bundle").disabled = false;
    }
}

async function handleReadAddToExisting(index) {
    const att = Office.context.mailbox.item.attachments[index];
    document.querySelectorAll(".btn-upload-single").forEach(b => { b.disabled = true; });
    showReadStatus("Signing in\u2026");

    try {
        const token = await getAuthToken();
        const latest = _getLatestNonShareRecord();
        if (!latest) throw new Error("No existing conversation found.");
        const { conversationId: existingConvId, documentId: existingDocId } = latest;

        showReadStatus("Adding " + att.name + " to conversation\u2026");
        await uploadSupportingById(att, existingConvId, token);

        if (_customProps) {
            saveConversationRecord(_customProps, singleFingerprint(att), {
                conversationId: existingConvId, documentId: existingDocId,
                label: att.name, uploadType: "bundle", timestamp: Date.now(),
            }).catch(err => console.warn("customProps save failed:", err.message));
        }

        showReadStatus("");
        await enterChat(existingConvId, existingDocId, token);
    } catch (err) {
        console.error("Add to existing error:", err);
        showReadStatus("Error: " + err.message);
        clearToken();
        document.querySelectorAll(".btn-upload-single").forEach(b => { b.disabled = false; });
    }
}

// ── Helpers ────────────────────────────────────────────────────────────────

function _getLatestNonShareRecord() {
    const cpR  = Object.values(getConversationMap(_customProps || {})).filter(r => r.uploadType !== "shared-link");
    const thR  = getThreadContextAll().filter(r => r.uploadType !== "shared-link");
    const all  = [...cpR, ...thR.filter(r => !cpR.some(c => c.conversationId === r.conversationId))];
    return all.sort((a, b) => (b.timestamp || 0) - (a.timestamp || 0))[0] || null;
}

// ── Read-mode initialisation ───────────────────────────────────────────────

export function initRead() {
    document.querySelector(".header-title").textContent = "View Document";

    // Wire up static buttons
    document.getElementById("btn-send").onclick = () => import("./chat.js").then(m => m.sendChatMessage());
    document.getElementById("user-input").addEventListener("keydown", e => {
        if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); import("./chat.js").then(m => m.sendChatMessage()); }
    });
    document.getElementById("btn-back").onclick          = () => import("./navigation.js").then(m => m.switchToReadView());
    document.getElementById("btn-upload-bundle").onclick = handleReadBundleUpload;
    document.getElementById("chk-bulk").onchange         = onReadToggleMode;

    document.getElementById("view-read-init").classList.remove("hidden");

    extractShareLinkFromBody()
        .then(shareInfo => {
            document.getElementById("view-read-init").classList.add("hidden");
            document.getElementById("view-read").classList.remove("hidden");

            if (shareInfo && (shareInfo.conversationId || shareInfo.shareId)) {
                setReadShareInfo(shareInfo);
                renderShareSection(shareInfo);
            }

            loadCustomProps()
                .then(cp => {
                    setCustomProps(cp);
                    renderPreviousChats();
                    loadReadAttachments();

                    const atts = Office.context.mailbox.item.attachments || [];
                    if (shareInfo?.conversationId && atts.length > 0)
                        document.getElementById("read-or-divider").classList.remove("hidden");

                    _autoOpenLatestConversation(shareInfo);
                })
                .catch(() => {
                    renderPreviousChats();
                    loadReadAttachments();
                    const atts = Office.context.mailbox.item.attachments || [];
                    if (shareInfo?.conversationId && atts.length > 0)
                        document.getElementById("read-or-divider").classList.remove("hidden");
                });
        })
        .catch(err => showReadInitError("Error reading email: " + err.message));
}

function onReadToggleMode() {
    const bulk = document.getElementById("chk-bulk").checked;
    document.getElementById("lbl-bundle").classList.toggle("active", bulk);
    document.getElementById("lbl-individual").classList.toggle("active", !bulk);
    document.getElementById("bundle-footer").classList.toggle("hidden", !bulk);
    loadReadAttachments();
}

async function _autoOpenLatestConversation(shareInfo) {
    const cpPrimary = Object.values(getConversationMap(_customProps))
        .filter(r => r.uploadType !== "shared-link");
    const thPrimary = getThreadContextAll().filter(r => r.uploadType !== "shared-link");
    const primary   = [...cpPrimary, ...thPrimary.filter(r => !cpPrimary.some(c => c.conversationId === r.conversationId))]
        .sort((a, b) => (b.timestamp || 0) - (a.timestamp || 0));

    try {
        const token = await getAuthToken();
        if (primary.length > 0) {
            const latest = primary[0];
            return await enterChat(latest.conversationId, latest.documentId, token);
        }
        if (shareInfo?.shareId) {
            const { conversationId, docId } = await resolveShareId(shareInfo.shareId, token);
            shareInfo.conversationId = conversationId;
            shareInfo.docId          = docId;
            return await enterChat(conversationId, docId, token);
        }
        if (shareInfo?.conversationId) {
            return await enterChat(shareInfo.conversationId, shareInfo.docId, token);
        }
    } catch (err) {
        console.warn("Auto-resume failed:", err.message);
    }
}
