export function parseDocUrl(rawUrl) {
    try {
        const url = rawUrl
            .replace(/[>)"'\s]+$/, "")     // strip trailing junk
            .replace(/&amp;/gi, "&");       // decode HTML entity Outlook encodes in href attributes
        const u  = new URL(url);
        const sp = u.searchParams;

        const conversationId =
            sp.get("conversation-id") ||   // hyphenated (proxy URL format)
            sp.get("conversation_id") ||
            sp.get("conversationId")  ||
            sp.get("cid")             || null;

        const docId =
            sp.get("doc-id")     ||        // hyphenated (proxy URL format)
            sp.get("doc_id")     ||
            sp.get("documentId") ||
            sp.get("did")        || null;

        if (conversationId) return { conversationId, docId, shareUrl: url };

        // Path-based fallback: /conversationId/docId
        const segments = u.pathname.split("/").filter(Boolean);
        if (segments.length >= 2)
            return { conversationId: segments[segments.length - 2], docId: segments[segments.length - 1], shareUrl: url };
        if (segments.length === 1)
            return { conversationId: segments[0], docId: null, shareUrl: url };
    } catch (_) {}
    return { conversationId: null, docId: null, shareUrl: null };
}

export function extractShareLinkFromBody() {
    return new Promise((resolve, reject) => {
        Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, (result) => {
            if (result.status !== Office.AsyncResultStatus.Succeeded) {
                reject(new Error(result.error?.message || "Body read failed"));
                return;
            }

            const html = result.value || "";

            // 1. Anchor hrefs with visible text (for display name)
            const anchorRe = /<a[^>]+href=["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>/gi;
            let m;
            while ((m = anchorRe.exec(html)) !== null) {
                const href     = m[1];
                const linkText = m[2].replace(/<[^>]+>/g, "").replace(/\s+/g, " ").trim();
                const parsed   = parseDocUrl(href);
                if (parsed.conversationId) {
                    resolve({ ...parsed, linkText: linkText || href });
                    return;
                }
            }

            // 2. Fallback: bare URLs
            const urlRe = /https?:\/\/[^\s"'<>)]+/gi;
            while ((m = urlRe.exec(html)) !== null) {
                const parsed = parseDocUrl(m[0]);
                if (parsed.conversationId) {
                    resolve({ ...parsed, linkText: null });
                    return;
                }
            }

            resolve({ conversationId: null, docId: null, linkText: null });
        });
    });
}
