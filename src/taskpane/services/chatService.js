import { WELCOME_URL, ASK_URL, CONVERSATION_URL } from "../config.js";

export async function fetchHistory(token, conversationId, documentId) {
    return fetch(
        `${CONVERSATION_URL}/history?conversation_id=${encodeURIComponent(conversationId)}&document_id=${encodeURIComponent(documentId)}`,
        {
            headers: {
                Authorization: "Bearer " + token,
                "ngrok-skip-browser-warning": "true",
            },
        }
    );
}

export async function fetchWelcome(token, conversationId, documentId) {
    return fetch(WELCOME_URL, {
        method: "POST",
        headers: {
            "Content-Type": "application/json",
            Authorization: "Bearer " + token,
            "ngrok-skip-browser-warning": "true",
        },
        body: JSON.stringify({ conversationId, documentId }),
    });
}

export async function askQuestion(token, conversationId, text) {
    return fetch(ASK_URL, {
        method: "POST",
        headers: { "Content-Type": "application/json", Authorization: "Bearer " + token },
        body:   JSON.stringify({ conversationId, text, isMobile: false }),
    });
}
