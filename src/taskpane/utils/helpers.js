// ── Pure Utilities ──────────────────────────────────────────────────────────
// No Office API, no DOM, no state imports — safe to unit-test in isolation.

import { BLUE_BASE, SMARTBLUE_DOMAINS } from "../config.js";

/**
 * Escape a string for safe HTML interpolation.
 */
export function escHtml(str) {
    return (str || "")
        .replace(/&/g, "&amp;").replace(/</g, "&lt;")
        .replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

/**
 * Format a byte count into a human-readable string (B / KB / MB).
 */
export function formatBytes(bytes) {
    if (!bytes) return "";
    if (bytes < 1024)    return bytes + " B";
    if (bytes < 1048576) return (bytes / 1024).toFixed(1) + " KB";
    return (bytes / 1048576).toFixed(1) + " MB";
}

/**
 * Convert SmartBlue API markdown-like response text into HTML.
 * Handles <blueEmbed-doc-page> citation tags, **bold**, bullet lists,
 * and plain paragraphs.
 */
export function formatResponse(raw, conversationId, documentId) {
    const base_url = `${BLUE_BASE}/conversation?conversation-id=${conversationId}&doc-id=${documentId}`;

    let text = raw.replace(
        /<blueEmbed-doc-page>([^<]+)<\/blueEmbed-doc-page>/g,
        (_, inner) => {
            const parts      = inner.trim().split(":");
            const embedDocId = parts[0] || "";
            const page       = parts[parts.length - 1] || "1";
            const fragment   = embedDocId && embedDocId !== documentId
                ? `#sub-doc-id=${encodeURIComponent(embedDocId)}&page=${page}`
                : `#page=${page}`;
            return `<a href="${base_url}${fragment}" target="_blank" class="page-ref" data-page="${page}">pg ${page}</a>`;
        }
    );

    text = text.replace(/\*\*(.+?)\*\*/g, "<strong>$1</strong>");

    const lines = text.split(/\n/);
    let html = "", inList = false;
    for (const rawLine of lines) {
        const line = rawLine.trim();
        if (!line) {
            if (inList) { html += "</ul>"; inList = false; }
            continue;
        }
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

/**
 * Return the MIME type for a filename based on its extension.
 */
export function getMimeType(filename) {
    const ext = (filename || "").split(".").pop().toLowerCase();
    const MAP = {
        pdf:  "application/pdf",
        doc:  "application/msword",
        docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        xls:  "application/vnd.ms-excel",
        xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        csv:  "text/csv",
        ppt:  "application/vnd.ms-powerpoint",
        pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        txt:  "text/plain",
        rtf:  "application/rtf",
        png:  "image/png",
        jpg:  "image/jpeg",
        jpeg: "image/jpeg",
        gif:  "image/gif",
        webp: "image/webp",
        zip:  "application/zip",
    };
    return MAP[ext] || "application/octet-stream";
}

/**
 * Clipboard copy with execCommand fallback for older browsers / Outlook.
 * @param {string}   text
 * @param {Function} cb   Called on success.
 */
export function fallbackCopy(text, cb) {
    const ta = document.createElement("textarea");
    ta.value = text;
    ta.style.cssText = "position:fixed;opacity:0";
    document.body.appendChild(ta);
    ta.select();
    try { document.execCommand("copy"); cb(); } catch (_) {}
    document.body.removeChild(ta);
}

// ── URL / share-link parsing ───────────────────────────────────────────────

/**
 * Parse a raw URL and extract SmartBlue conversation/doc/share IDs.
 * Returns { conversationId, docId, shareId, shareUrl } — nulls when not found.
 */
export function parseDocUrl(rawUrl) {
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
        if (shareId)         return { conversationId: null, docId: null, shareId, shareUrl: url };
        if (conversationId)  return { conversationId, docId, shareUrl: url };
    } catch (_) {}
    return { conversationId: null, docId: null, shareId: null, shareUrl: null };
}
