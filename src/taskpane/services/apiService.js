// ── API Service ─────────────────────────────────────────────────────────────
// All HTTP calls to the SmartBlue / proxy backend.

import {
    PROXY_BASE, BLUE_BASE,
    SHARE_URL, WELCOME_URL, ASK_URL, CONVERSATION_URL,
} from "../config.js";

const NGROK_HEADER = { "ngrok-skip-browser-warning": "true" };

// ── Sharing ────────────────────────────────────────────────────────────────

/**
 * Notify the backend of recipients who should receive access.
 * Returns the share URL string.
 */
export async function callShareApi(token, conversationId, docId, senderEmail, recipients) {
    const payload = { sender_email: senderEmail, recipients };
    if (conversationId) payload.conversation_id = conversationId;
    if (docId)          payload.doc_id           = docId;

    const resp = await fetch(SHARE_URL, {
        method:  "POST",
        headers: { "Content-Type": "application/json", Authorization: "Bearer " + token },
        body:    JSON.stringify(payload),
    });
    if (!resp.ok) throw new Error("Share API failed (" + resp.status + "): " + await resp.text());

    const data = await resp.json();
    const url  = data.share_url || data.shareUrl || data.url || "";
    if (!url) throw new Error("Share API returned no URL.");
    return url;
}

/**
 * Resolve a share-id token to { conversationId, docId }.
 */
export async function resolveShareId(shareId, token) {
    const resp = await fetch(
        `${PROXY_BASE}/v1/conversation/share/${encodeURIComponent(shareId)}`,
        { headers: { Authorization: "Bearer " + token, ...NGROK_HEADER } }
    );
    if (!resp.ok) throw new Error("Share resolve failed (" + resp.status + "): " + await resp.text());

    const data = await resp.json();
    const conversationId = data.conversation_id || data.conversationId || null;
    const docId          = data.doc_id          || data.docId          || null;
    if (!conversationId) throw new Error("No conversation_id returned from share resolve.");
    if (!docId)          throw new Error("No doc_id returned from share resolve.");
    return { conversationId, docId };
}

/**
 * Create (or refresh) a shareable link for a document.
 * Fetches the current recipient list first, then re-posts with the same list
 * so existing access is preserved.
 * Returns the share URL string.
 */
export async function getShareLink(token, conversationId, docId, access = "restricted") {
    // Step 1: fetch current recipient list
    const emailList = await fetch(
        `${PROXY_BASE}/v1/doc-access/share/${encodeURIComponent(docId)}/list`,
        { headers: { Authorization: "Bearer " + token, ...NGROK_HEADER } }
    )
        .then(resp => {
            if (!resp.ok) throw new Error("Recipient list fetch failed (" + resp.status + ")");
            return resp.json();
        })
        .then(data => (data.users || []).map(u => u.email).filter(Boolean))
        .catch(err => { console.warn("Could not fetch recipient list:", err.message); return []; });

    // Step 2: POST to share endpoint
    const resp = await fetch(
        `${PROXY_BASE}/v1/document/${encodeURIComponent(docId)}/share`,
        {
            method:  "POST",
            headers: {
                Authorization:  "Bearer " + token,
                "Content-Type": "application/json",
                ...NGROK_HEADER,
            },
            body: JSON.stringify({
                receivers:       emailList.map(email => ({ email })),
                allowed_domains: ["*"],
                roles_allowed:   [access],
                expire_in_secs:  30 * 24 * 60 * 60,
                allow_download:  false,
                allow_handsfree: false,
                text_notes:      [""],
                voice_notes:     [],
            }),
        }
    );
    if (!resp.ok) throw new Error("Share link API failed (" + resp.status + "): " + await resp.text());

    const data = await resp.json();
    const url  = data["share-url"] || "";
    if (!url) throw new Error("Share link API returned no URL.");
    return url;
}

// ── Conversation ───────────────────────────────────────────────────────────

export function fetchHistory(token, conversationId) {
    return fetch(
        `${CONVERSATION_URL}/history?conversation_id=${encodeURIComponent(conversationId)}`,
        { headers: { Authorization: "Bearer " + token, ...NGROK_HEADER } }
    );
}

export function fetchWelcome(token, conversationId, documentId) {
    return fetch(WELCOME_URL, {
        method:  "POST",
        headers: {
            "Content-Type": "application/json",
            Authorization:  "Bearer " + token,
            ...NGROK_HEADER,
        },
        body: JSON.stringify({ conversationId, documentId }),
    });
}

export function askQuestion(token, conversationId, text) {
    return fetch(ASK_URL, {
        method:  "POST",
        headers: {
            "Content-Type": "application/json",
            Authorization:  "Bearer " + token,
        },
        body: JSON.stringify({ conversationId, text, isMobile: false }),
    });
}
