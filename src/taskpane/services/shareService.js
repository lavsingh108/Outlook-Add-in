import { SHARE_URL } from "../config.js";

export async function callShareApi(token, conversationId, docId, senderEmail, recipients) {
    const payload = { sender_email: senderEmail, recipients };
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
