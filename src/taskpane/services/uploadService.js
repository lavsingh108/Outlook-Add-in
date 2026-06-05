// ── Upload Service ──────────────────────────────────────────────────────────
// Fetches attachments from the Office mailbox and uploads them to the proxy.

import { UPLOAD_URL, BUNDLE_ADD_URL } from "../config.js";
import { getMimeType } from "../utils/helpers.js";

/**
 * Read an Outlook attachment and return a Blob.
 */
export function getAttachmentBlob(attachmentId, filename) {
    return new Promise((resolve, reject) => {
        Office.context.mailbox.item.getAttachmentContentAsync(attachmentId, result => {
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

/**
 * Upload a primary attachment, creating a new conversation.
 * Returns { conversationId, documentId }.
 */
export async function uploadPrimary(att, token) {
    const blob = await getAttachmentBlob(att.id, att.name);
    const form = new FormData();
    form.append("document", blob, att.name);

    const resp = await fetch(UPLOAD_URL, {
        method:  "POST",
        headers: { Authorization: "Bearer " + token },
        body:    form,
    });
    if (!resp.ok) throw new Error("Upload failed (" + resp.status + "): " + await resp.text());

    const data = await resp.json();
    console.log("Upload response:", JSON.stringify(data));

    const conversationId = data.conversation_id || data.conversationId || null;
    const documentId     = data.doc_id || data.document_id || data.documentId ||
                           data.docId  || data.doc         || data.id          || null;

    if (!conversationId) throw new Error("No conversation_id returned by upload.");
    if (!documentId)     throw new Error("No document ID returned by upload (keys: " + Object.keys(data).join(", ") + ").");

    return { conversationId, documentId };
}

/**
 * Upload a supporting attachment into an existing conversation (bundle add).
 * Logs a warning on failure — non-fatal by design.
 */
export async function uploadSupportingById(att, conversationId, token) {
    const blob = await getAttachmentBlob(att.id, att.name);
    const form = new FormData();
    form.append("document", blob, att.name);

    const resp = await fetch(
        `${BUNDLE_ADD_URL}?conversation_id=${encodeURIComponent(conversationId)}`,
        { method: "POST", headers: { Authorization: "Bearer " + token }, body: form }
    );
    if (!resp.ok) console.warn("Supporting upload failed:", att.name, await resp.text());
}

/**
 * Remove an attachment from the compose item if the "include attachment"
 * checkbox is checked. Accepts a single id or an array of ids.
 * Always resolves — user-added attachments fail silently (non-fatal).
 */
export async function removeAttachmentIfRequested(attachmentIds) {
    const checkbox = document.getElementById("chk-include-attachment");
    if (!checkbox || !checkbox.checked) return;

    const ids = Array.isArray(attachmentIds) ? attachmentIds : [attachmentIds];
    for (const id of ids) {
        await new Promise(resolve => {
            Office.context.mailbox.item.removeAttachmentAsync(id, result => {
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
