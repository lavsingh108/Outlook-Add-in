import { formatBytes } from "../utils/formatUtils.js";
import { escHtml } from "../utils/domUtils.js";

// Shared bundle list renderer — used by both read and compose bundle modes
export function renderBundleList(attachments, container) {
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
                    <div class="att-name" title="${escHtml(att.name)}">${escHtml(att.name)}</div>
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

export function updateBundleSelection(container) {
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

// onUpload(index) — called when user clicks Upload on an individual attachment
export function renderIndividualReadList(attachments, container, onUpload) {
    container.innerHTML = "";
    attachments.forEach((att, index) => {
        const div = document.createElement("div");
        div.className = "att-item";
        div.innerHTML = `
            <div class="att-individual-row">
                <div class="att-info">
                    <div class="att-name" title="${escHtml(att.name)}">${escHtml(att.name)}</div>
                    <div class="att-meta">${formatBytes(att.size)}</div>
                </div>
                <button class="btn-upload-single" data-index="${index}">Upload</button>
            </div>`;
        container.appendChild(div);
    });
    container.querySelectorAll(".btn-upload-single").forEach(btn => {
        btn.onclick = () => onUpload(parseInt(btn.dataset.index));
    });
}

// onUpload(index) — called when user clicks Share on an individual attachment
export function renderIndividualComposeList(attachments, container, onUpload) {
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
        btn.onclick = () => onUpload(parseInt(btn.dataset.index));
    });
}
