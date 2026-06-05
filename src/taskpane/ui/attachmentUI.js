// ── Attachment UI ────────────────────────────────────────────────────────────
// Renders attachment lists for both read mode (upload / bundle) and compose
// mode (share / add-to-bundle).

import { escHtml, formatBytes } from "../utils/helpers.js";
import { getAttachmentRecord, isAttachmentUploaded } from "../services/storageService.js";
import { _customProps, _composeConversationCtx, _composeUploadedAttIds } from "../state.js";
import { getAuthToken, clearToken } from "../services/authService.js";
import { showReadStatus } from "./chatUI.js";

// ── Read-mode: bundle (radio + checkboxes) ─────────────────────────────────

/**
 * Render the primary/secondary radio+checkbox list used in bundle upload mode.
 * @param {object[]} attachments
 * @param {HTMLElement} container
 * @param {Function} onEnterChat  (rec) => Promise
 * @param {Function} onBundleUpload  () => void — called when footer button triggers upload
 */
export function renderBundleList(attachments, container, onEnterChat, onBundleUpload) {
    container.innerHTML = "";

    attachments.forEach((att, index) => {
        const isPrimary = index === 0;
        const div = document.createElement("div");
        div.className  = "att-item" + (isPrimary ? " is-primary" : "");
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
                    ${isAttachmentUploaded(att, _customProps) && !isPrimary
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
            const atts = Office.context.mailbox.item.attachments || attachments;
            const primaryAtt = atts[parseInt(radio.value)];
            if (!primaryAtt) return;

            const bundleBtn = document.getElementById("btn-upload-bundle");
            if (!bundleBtn) return;

            const rec = getAttachmentRecord(primaryAtt, _customProps);
            if (rec) {
                bundleBtn.textContent = "\u25B6 Start Chat";
                bundleBtn.classList.add("btn-start-chat-att");
                bundleBtn.onclick = async () => {
                    bundleBtn.disabled = true;
                    showReadStatus("Signing in\u2026");
                    try {
                        const token = await getAuthToken();
                        await onEnterChat(rec.conversationId, rec.documentId, token);
                    } catch (err) {
                        showReadStatus("Error: " + err.message);
                        clearToken();
                        bundleBtn.disabled = false;
                    }
                };
            } else {
                bundleBtn.textContent = "\u2B06 Upload & Analyse";
                bundleBtn.classList.remove("btn-start-chat-att");
                bundleBtn.onclick = onBundleUpload;
            }
        });
    });
}

export function updateBundleSelection(container) {
    const primaryVal = container.querySelector("input[name='primaryIndex']:checked")?.value;
    container.querySelectorAll(".att-item").forEach(item => {
        const idx       = item.dataset.index;
        const isPrimary = idx === primaryVal;
        const secChk    = item.querySelector("input[name='secondaryIndex']");
        item.classList.toggle("is-primary", isPrimary);
        if (secChk) {
            if (isPrimary) { secChk.checked = false; secChk.disabled = true; }
            else { secChk.disabled = false; if (!secChk.dataset.userUnchecked) secChk.checked = true; }
        }
    });
}

// Keep userUnchecked flag in sync when the user manually unticks a secondary
document.addEventListener("change", e => {
    if (e.target.name === "secondaryIndex") {
        e.target.dataset.userUnchecked = e.target.checked ? "" : "1";
    }
});

// ── Read-mode: individual (one Upload button per attachment) ───────────────

/**
 * @param {object[]} attachments
 * @param {HTMLElement} container
 * @param {boolean} hasContext  — true when an existing conversation is available
 * @param {Function} onEnterChat
 * @param {Function} onSingleUpload  (index) => void
 * @param {Function} onAddToExisting (index) => void
 */
export function renderIndividualReadList(
    attachments, container, hasContext,
    onEnterChat, onSingleUpload, onAddToExisting
) {
    container.innerHTML = "";

    attachments.forEach((att, index) => {
        const div = document.createElement("div");
        div.className = "att-item";
        const rec      = getAttachmentRecord(att, _customProps);
        const btnLabel = rec ? "\u25B6 Start Chat" : hasContext ? "Add to Bundle" : "Upload";
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
        const r = getAttachmentRecord(Office.context.mailbox.item.attachments[i], _customProps);
        if (r) {
            btn.onclick = async () => {
                btn.disabled = true;
                showReadStatus("Signing in\u2026");
                try {
                    const token = await getAuthToken();
                    await onEnterChat(r.conversationId, r.documentId, token);
                } catch (err) {
                    showReadStatus("Error: " + err.message);
                    clearToken();
                    btn.disabled = false;
                }
            };
        } else {
            btn.onclick = hasContext ? () => onAddToExisting(i) : () => onSingleUpload(i);
        }
    });
}

// ── Read-mode: add-to-bundle (all are supporting, no radio) ───────────────

export function renderAddToBundleList(attachments, container) {
    container.innerHTML = "";
    attachments.forEach((att, index) => {
        const div = document.createElement("div");
        div.className   = "att-item";
        div.dataset.index = index;
        const alreadyAdded = isAttachmentUploaded(att, _customProps);
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

// ── Compose-mode: individual share buttons ─────────────────────────────────

/**
 * @param {object[]} attachments
 * @param {HTMLElement} container
 * @param {Function} onSingleUpload  (index) => void
 * @param {Function} onAddToBundle   (index) => void
 */
export function renderIndividualComposeList(
    attachments, container, onSingleUpload, onAddToBundle
) {
    attachments.forEach((att, index) => {
        const div = document.createElement("div");
        div.className = "att-item";

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
                    ${alreadyUploaded ? "disabled" : ""}>${btnLabel}</button>
            </div>`;
        container.appendChild(div);
    });

    container.querySelectorAll(".btn-upload-share:not([disabled])").forEach(btn => {
        const i = parseInt(btn.dataset.index);
        btn.onclick = _composeConversationCtx
            ? () => onAddToBundle(i)
            : () => onSingleUpload(i);
    });
}
