// ── Compose Mode ─────────────────────────────────────────────────────────────
// Handles the compose (Share Document) pane: recipient/attachment display,
// upload and share flows, and post-upload actions.

import { escHtml, fallbackCopy } from "./utils/helpers.js";
import { BLUE_BASE } from "./config.js";
import { getAuthToken, clearToken } from "./services/authService.js";
import {
    saveConversationRecord, getThreadContextAll, saveThreadContext,
    singleFingerprint, loadComposePrefs, saveComposePrefs,
} from "./services/storageService.js";
import {
    uploadPrimary, uploadSupportingById, removeAttachmentIfRequested,
} from "./services/uploadService.js";
import { callShareApi, getShareLink } from "./services/apiService.js";
import {
    renderBundleList, renderIndividualComposeList,
} from "./ui/attachmentUI.js";
import { showComposeStatus } from "./ui/chatUI.js";
import {
    state,
    _customProps, _composeAttachments, _composeRecipients,
    _senderEmail, _composeConversationCtx, _composeUploadedAttIds,
    _composeSharedRecipients, _composeRefreshTimer, _composeAccessLevel,
    setComposeAttachments, setComposeRecipients, setSenderEmail,
    setComposeConversationCtx, setComposeUploadedAttIds,
    setComposeSharedRecipients, setComposeRefreshTimer, setComposeAccessLevel,
} from "./state.js";

// ── Initialisation ─────────────────────────────────────────────────────────

export function initCompose() {
    document.querySelector(".header-title").textContent = "Share Document";
    document.getElementById("view-compose").classList.remove("hidden");
    document.getElementById("btn-refresh").classList.remove("hidden");

    document.getElementById("btn-refresh").onclick = () => {
        clearTimeout(_composeRefreshTimer);
        state.suppressAttachmentRefresh = false;
        setComposeConversationCtx(null);
        setComposeUploadedAttIds(new Set());
        setComposeSharedRecipients(new Set());

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

    document.getElementById("btn-back").onclick = () => import("./navigation.js").then(m => m.switchToReadView());
    document.getElementById("btn-send").onclick  = () => import("./chat.js").then(m => m.sendChatMessage());
    document.getElementById("user-input").addEventListener("keydown", e => {
        if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); import("./chat.js").then(m => m.sendChatMessage()); }
    });

    document.getElementById("sel-access").value    = _composeAccessLevel;
    document.getElementById("sel-access").onchange = () => {
        setComposeAccessLevel(document.getElementById("sel-access").value);
        saveComposePrefs();
    };
    document.getElementById("chk-include-attachment").onchange = () => saveComposePrefs();

    loadComposePrefs();

    document.getElementById("btn-compose-upload").onclick = handleComposeBundleUpload;
    document.getElementById("btn-copy-link").onclick      = copyResultLink;
    document.getElementById("chk-compose-bulk").onchange  = onComposeToggleMode;

    loadComposeData(false);

    // Live sync — debounce attachment/recipient change events
    const debouncedRefresh = () => {
        if (state.suppressAttachmentRefresh) return;
        clearTimeout(_composeRefreshTimer);
        setComposeRefreshTimer(setTimeout(() => loadComposeData(true), 400));
    };
    if (Office.context.requirements.isSetSupported("Mailbox", "1.8")) {
        Office.context.mailbox.item.addHandlerAsync(Office.EventType.AttachmentsChanged, debouncedRefresh);
    }
    if (Office.context.requirements.isSetSupported("Mailbox", "1.7")) {
        Office.context.mailbox.item.addHandlerAsync(Office.EventType.RecipientsChanged, debouncedRefresh);
    }
}

// ── Data loading ───────────────────────────────────────────────────────────

function isComposeBulkMode() {
    return document.getElementById("chk-compose-bulk").checked;
}

function onComposeToggleMode() {
    const bulk = isComposeBulkMode();
    document.getElementById("clbl-bundle")?.classList.toggle("active", bulk);
    document.getElementById("clbl-individual")?.classList.toggle("active", !bulk);
    document.getElementById("compose-bundle-footer").classList.toggle("hidden", !bulk);
    renderComposeAttachments(_composeAttachments);
}

export function loadComposeData(isRefresh) {
    if (state.suppressAttachmentRefresh) {
        document.getElementById("btn-refresh").classList.remove("spinning");
        return;
    }
    if (isRefresh) {
        document.getElementById("btn-refresh").classList.add("spinning");
        if (!_composeConversationCtx) {
            document.getElementById("compose-result").classList.add("hidden");
        }
        showComposeStatus("");
    }

    try { setSenderEmail(Office.context.mailbox.userProfile.emailAddress || ""); }
    catch (e) { setSenderEmail(""); }

    const item = Office.context.mailbox.item;
    Promise.all([
        new Promise(res => item.to.getAsync(r  => res(r.status === Office.AsyncResultStatus.Succeeded ? r.value : []))),
        new Promise(res => item.cc.getAsync(r  => res(r.status === Office.AsyncResultStatus.Succeeded ? r.value : []))),
        new Promise(res => item.bcc.getAsync(r => res(r.status === Office.AsyncResultStatus.Succeeded ? r.value : []))),
        new Promise(res => item.getAttachmentsAsync(r => res(r.status === Office.AsyncResultStatus.Succeeded ? r.value : []))),
    ]).then(([toList, ccList, bccList, attachments]) => {
        const seen = new Set();
        setComposeRecipients(
            [...toList, ...ccList, ...bccList]
                .map(r => (r.emailAddress || "").toLowerCase().trim())
                .filter(e => { if (!e || seen.has(e)) return false; seen.add(e); return true; })
        );
        renderComposeRecipients(toList, ccList, bccList);
        setComposeAttachments(attachments);

        // Auto-select mode based on attachment count
        const bulkChk = document.getElementById("chk-compose-bulk");
        if (attachments.length === 1)      bulkChk.checked = false;
        else if (attachments.length > 1)   bulkChk.checked = true;
        onComposeToggleMode();

        renderComposeAttachments(attachments);
        document.getElementById("btn-refresh").classList.remove("spinning");
    });
}

// ── Recipients rendering ───────────────────────────────────────────────────

function renderComposeRecipients(toList, ccList, bccList = []) {
    const area  = document.getElementById("compose-recipients");
    const badge = document.getElementById("recipients-count");
    const total = toList.length + ccList.length + bccList.length;
    badge.textContent = total || "";

    if (total === 0) {
        area.innerHTML = `<div class="compose-empty">No recipients yet. Add To / CC addresses then click &#8635; Refresh.</div>`;
        return;
    }
    area.innerHTML = "";

    const buildRow = (label, list) => {
        if (!list.length) return;
        const row   = document.createElement("div"); row.className = "recipient-row";
        const lbl   = document.createElement("span"); lbl.className = "recipient-row-label"; lbl.textContent = label;
        row.appendChild(lbl);
        const chips = document.createElement("div"); chips.className = "recipient-chips";
        list.forEach(r => {
            const chip = document.createElement("span"); chip.className = "recipient-chip";
            chip.textContent = r.emailAddress || r.displayName || "";
            chip.title       = r.emailAddress || "";
            chips.appendChild(chip);
        });
        row.appendChild(chips);
        area.appendChild(row);
    };
    buildRow("To:", toList); buildRow("CC:", ccList); buildRow("BCC:", bccList);
}

// ── Attachments rendering ──────────────────────────────────────────────────

function renderComposeAttachments(attachments) {
    const list       = document.getElementById("compose-attachments");
    const badge      = document.getElementById("attachments-count");
    const bulkSwitch = document.getElementById("chk-compose-bulk");
    badge.textContent = attachments.length || "";

    if (!attachments.length) {
        list.innerHTML = `<div class="compose-empty">No attachments yet. Attach a document then click &#8635; Refresh.</div>`;
        document.getElementById("compose-bundle-footer").classList.add("hidden");
        document.getElementById("compose-attachment-option").classList.add("hidden");
        document.getElementById("compose-access-row")?.classList.add("hidden");
        document.getElementById("btn-compose-upload").disabled = true;
        if (_composeConversationCtx) renderPostUploadActions();
        return;
    }

    document.getElementById("compose-attachment-option").classList.remove("hidden");
    if (!_composeConversationCtx) document.getElementById("compose-access-row")?.classList.remove("hidden");

    if (attachments.length === 1) { bulkSwitch.checked = false; bulkSwitch.disabled = true; }
    else                          { bulkSwitch.disabled = false; }

    list.innerHTML = "";

    // After an upload: always individual so each attachment shows its own state
    if (_composeConversationCtx) {
        document.getElementById("compose-bundle-footer").classList.add("hidden");
        document.getElementById("chk-compose-bulk").closest("label")?.classList.add("hidden");
        document.getElementById("clbl-bundle")?.classList.add("hidden");
        document.getElementById("clbl-individual")?.classList.add("hidden");

        const newAtts     = attachments.filter(a => !_composeUploadedAttIds.has(a.id));
        const docsSection = document.getElementById("compose-documents-section");
        if (newAtts.length === 0) {
            docsSection?.classList.add("hidden");
            document.getElementById("compose-attachment-option").classList.add("hidden");
        } else {
            docsSection?.classList.remove("hidden");
            badge.textContent = newAtts.length;
            renderIndividualComposeList(newAtts, list, handleComposeSingleUpload, handleComposeAddToBundle);
        }
        renderPostUploadActions();
        return;
    }

    // Normal pre-upload rendering
    if (isComposeBulkMode()) {
        document.getElementById("compose-bundle-footer").classList.remove("hidden");
        document.getElementById("btn-compose-upload").disabled = false;
        renderBundleList(attachments, list, () => {}, handleComposeBundleUpload);
    } else {
        document.getElementById("compose-bundle-footer").classList.add("hidden");
        renderIndividualComposeList(attachments, list, handleComposeSingleUpload, handleComposeAddToBundle);
    }
}

// ── Post-upload: new-recipients section ───────────────────────────────────

function renderPostUploadActions() {
    const newRecipients = _composeRecipients.filter(e => !_composeSharedRecipients.has(e));
    let section = document.getElementById("compose-new-recipients");

    if (newRecipients.length > 0) {
        if (!section) {
            section = document.createElement("div");
            section.id        = "compose-new-recipients";
            section.className = "compose-section";
            const recipSection = document.getElementById("compose-recipients")?.closest(".compose-section");
            if (recipSection) recipSection.insertAdjacentElement("afterend", section);
            else document.getElementById("compose-body")?.appendChild(section);
        }
        section.innerHTML = `
            <div class="compose-section-header">
                <span class="compose-section-title" style="font-size:11px;color:#5f6b7a">New Recipients</span>
                <span class="compose-badge">${newRecipients.length}</span>
                <button class="btn-share-new-recip">Share</button>
            </div>
            <div class="recipient-chips" style="flex-wrap:wrap;gap:4px;margin-top:4px">
                ${newRecipients.map(e => `<span class="recipient-chip">${escHtml(e)}</span>`).join("")}
            </div>`;

        section.querySelector(".btn-share-new-recip").onclick = async () => {
            const btn = section.querySelector(".btn-share-new-recip");
            btn.disabled = true; btn.textContent = "Sharing\u2026";
            try {
                const token = await getAuthToken();
                await callShareApi(
                    token, _composeConversationCtx.conversationId,
                    _composeConversationCtx.documentId, _senderEmail, newRecipients
                );
                newRecipients.forEach(e => _composeSharedRecipients.add(e));
                section.remove();
                showComposeStatus("\u2713 Shared with new recipients");
                setTimeout(() => showComposeStatus(""), 3000);
            } catch (err) {
                showComposeStatus("Share failed: " + err.message); clearToken();
                btn.disabled = false; btn.textContent = "Share with New Recipients";
            }
        };
    } else if (section) {
        section.remove();
    }
}

// ── Upload handlers ────────────────────────────────────────────────────────

async function handleComposeBundleUpload() {
    const primaryRadio = document.querySelector("input[name='primaryIndex']:checked");
    if (!primaryRadio) { showComposeStatus("Please select a primary document."); return; }

    const primaryIndex     = parseInt(primaryRadio.value);
    const primaryAtt       = _composeAttachments[primaryIndex];
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
            for (const idx of secondaryIndices)
                await uploadSupportingById(_composeAttachments[idx], conversationId, token);
        }

        showComposeStatus("Creating share link\u2026");
        await callShareApi(token, conversationId, documentId, _senderEmail, _composeRecipients);

        showComposeStatus("Inserting link into email\u2026");
        const accessLevel = document.getElementById("sel-access")?.value || _composeAccessLevel;
        const shareLink   = await getShareLink(token, conversationId, documentId, accessLevel);
        await insertShareLinkIntoBody(shareLink, primaryAtt.name);

        const allAttIds = [primaryAtt.id, ...secondaryIndices.map(i => _composeAttachments[i].id)];
        setComposeConversationCtx({ conversationId, documentId });
        allAttIds.forEach(id => _composeUploadedAttIds.add(id));
        _composeRecipients.forEach(e => _composeSharedRecipients.add(e));

        state.suppressAttachmentRefresh = true;
        await removeAttachmentIfRequested(allAttIds);
        setTimeout(() => { state.suppressAttachmentRefresh = false; loadComposeData(true); }, 1500);

        if (_customProps) {
            saveConversationRecord(_customProps, `compose_${conversationId}`, {
                conversationId, documentId, label: primaryAtt.name, uploadType: "bundle", timestamp: Date.now(),
            }).catch(err => console.warn("customProps save failed:", err.message));
        }
        saveThreadContext({ conversationId, documentId, label: primaryAtt.name, uploadType: "bundle", timestamp: Date.now() });

        document.getElementById("compose-bundle-footer").classList.add("hidden");
        showComposeStatus("");
        renderComposeResult(shareLink);
    } catch (err) {
        console.error("Compose bundle upload error:", err);
        showComposeStatus("Error: " + err.message);
        clearToken();
    } finally {
        uploadBtn.disabled = false;
    }
}

async function handleComposeAddToBundle(index) {
    if (!_composeConversationCtx) return;

    const btn   = document.querySelector(`.btn-upload-share[data-index="${index}"]`);
    const attId = btn?.dataset.attId;
    const att   = attId ? _composeAttachments.find(a => a.id === attId) : _composeAttachments[index];
    if (!att) return;

    document.querySelectorAll(".btn-upload-share").forEach(b => { b.disabled = true; });
    showComposeStatus("Adding " + att.name + " to bundle\u2026");

    try {
        const token = await getAuthToken();
        await uploadSupportingById(att, _composeConversationCtx.conversationId, token);

        if (_customProps) {
            saveConversationRecord(_customProps, singleFingerprint(att), {
                conversationId: _composeConversationCtx.conversationId,
                documentId:     _composeConversationCtx.documentId,
                label:          att.name, uploadType: "bundle", timestamp: Date.now(),
            }).catch(err => console.warn("customProps save failed:", err.message));
        }

        _composeUploadedAttIds.add(att.id);
        state.suppressAttachmentRefresh = true;
        await removeAttachmentIfRequested(_composeAttachments.map(a => a.id));
        setTimeout(() => { state.suppressAttachmentRefresh = false; }, 1500);

        btn?.closest(".att-item")?.remove();
        document.getElementById("compose-documents-section")?.classList.add("hidden");
        document.getElementById("compose-attachment-option")?.classList.add("hidden");

        showComposeStatus("\u2713 Added to bundle");
        setTimeout(() => showComposeStatus(""), 3000);
    } catch (err) {
        console.error("Compose add to bundle error:", err);
        showComposeStatus("Error: " + err.message);
        clearToken();
        document.querySelectorAll(".btn-upload-share").forEach(b => { b.disabled = false; });
    }
}

async function handleComposeSingleUpload(index) {
    const att = _composeAttachments[index];
    document.querySelectorAll(".btn-upload-share").forEach(b => { b.disabled = true; });
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
        const accessLevel = document.getElementById("sel-access")?.value || _composeAccessLevel;
        const shareLink   = await getShareLink(token, conversationId, documentId, accessLevel);
        await insertShareLinkIntoBody(shareLink, att.name);

        state.suppressAttachmentRefresh = true;
        await removeAttachmentIfRequested([att.id]);
        setTimeout(() => { state.suppressAttachmentRefresh = false; loadComposeData(true); }, 1500);

        if (_customProps) {
            saveConversationRecord(_customProps, `compose_${conversationId}`, {
                conversationId, documentId, label: att.name, uploadType: "single", timestamp: Date.now(),
            }).catch(err => console.warn("customProps save failed:", err.message));
        }
        saveThreadContext({ conversationId, documentId, label: att.name, uploadType: "single", timestamp: Date.now() });

        setComposeConversationCtx({ conversationId, documentId });
        _composeUploadedAttIds.add(att.id);
        _composeRecipients.forEach(e => _composeSharedRecipients.add(e));

        succeeded = true;
        showComposeStatus("");
        renderComposeResult(shareLink);
        renderComposeAttachments(_composeAttachments);
    } catch (err) {
        console.error("Compose single upload error:", err);
        showComposeStatus("Error: " + err.message);
        clearToken();
    } finally {
        document.querySelectorAll(".btn-upload-share").forEach(b => {
            const btnIndex = parseInt(b.dataset.index);
            if (succeeded && btnIndex === index) {
                b.textContent = "\u2713 Shared";
                b.disabled    = true;
                b.classList.add("btn-shared-done");
            } else {
                b.disabled = false;
            }
        });
    }
}

// ── Compose result UI ──────────────────────────────────────────────────────

function renderComposeResult(link) {
    document.getElementById("result-link-text").value = link;

    const copyBtn = document.getElementById("btn-copy-link");
    if (copyBtn) {
        copyBtn.onclick = () => {
            const restore = () => setTimeout(() => {
                copyBtn.innerHTML = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" width="14" height="14"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg> Copy URL';
            }, 1600);
            const done = () => { copyBtn.textContent = "\u2713 Copied!"; restore(); };
            if (navigator.clipboard) navigator.clipboard.writeText(link).then(done).catch(() => fallbackCopy(link, done));
            else fallbackCopy(link, done);
        };
    }

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
                const { setChatFromCompose } = await import("./state.js");
                setChatFromCompose(true);
                const { enterChat } = await import("./chat.js");
                await enterChat(_composeConversationCtx.conversationId, _composeConversationCtx.documentId, token);
            } catch (err) {
                showComposeStatus("Error: " + err.message);
                clearToken();
                startChatBtn.disabled = false;
            }
        };
    }

    document.getElementById("compose-access-row")?.classList.add("hidden");
    document.getElementById("compose-result").classList.remove("hidden");
    document.getElementById("compose-result").scrollIntoView({ behavior: "smooth" });
}

function copyResultLink() {
    const link = document.getElementById("result-link-text")?.value || "";
    if (link) {
        if (navigator.clipboard) navigator.clipboard.writeText(link).catch(() => fallbackCopy(link, () => {}));
        else fallbackCopy(link, () => {});
    }
}

function insertShareLinkIntoBody(link, filename) {
    return new Promise(resolve => {
        const html = `<p style="font-family:sans-serif;margin:8px 0;">`
            + `<a href="${link}" target="_blank" style="color:#0D47A1;font-size:14px;font-weight:500;text-decoration:none;">`
            + `${filename} \u2014 View on SmartBlue</a></p>`
            + `<p style="font-family:sans-serif;margin:8px 0;">Access the document directly in BlueAI using the link below for a secure and seamless viewing experience.</p>`;
        Office.context.mailbox.item.body.setSelectedDataAsync(
            html, { coercionType: Office.CoercionType.Html }, result => {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    Office.context.mailbox.item.body.setSelectedDataAsync(
                        `\n${filename}: ${link}\n`, { coercionType: Office.CoercionType.Text }, () => resolve()
                    );
                } else {
                    resolve();
                }
            }
        );
    });
}
