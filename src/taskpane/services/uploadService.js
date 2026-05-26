import { UPLOAD_URL, BUNDLE_ADD_URL } from "../config.js";
import { getMimeType } from "../utils/mimeUtils.js";

export function getAttachmentBlob(attachmentId, filename) {
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

export async function uploadPrimary(att, token) {
    const blob = await getAttachmentBlob(att.id, att.name);
    const form = new FormData();
    form.append("document", blob, att.name);

    const resp = await fetch(UPLOAD_URL, {
        method: "POST", headers: { Authorization: "Bearer " + token }, body: form,
    });
    if (!resp.ok) throw new Error("Upload failed (" + resp.status + "): " + await resp.text());

    const data = await resp.json();
    console.log("Upload response:", JSON.stringify(data));

    const conversationId = data.conversation_id || data.conversationId || null;
    const documentId =
        data.doc_id      ||
        data.document_id ||
        data.documentId  ||
        data.docId       ||
        data.doc         ||
        data.id          || null;

    if (!conversationId) throw new Error("No conversation_id returned by upload.");
    if (!documentId)     throw new Error("No document ID returned by upload (keys: " + Object.keys(data).join(", ") + ").");
    return { conversationId, documentId };
}

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

export const uploadSupporting = uploadSupportingById;
