import { escHtml } from "../utils/domUtils.js";

// onStartChat(shareInfo): async callback — should throw on error so the button is re-enabled
export function renderShareSection(shareInfo, onStartChat) {
    const section = document.getElementById("read-share-section");
    const card    = document.getElementById("read-share-card");

    const displayText = shareInfo.linkText || shareInfo.shareUrl || "View on SmartBlue";
    const displayUrl  = shareInfo.shareUrl  || "";

    card.innerHTML = `
        <div class="read-share-inner">
            <svg class="read-share-file-icon" viewBox="0 0 24 24" fill="none"
                 stroke="currentColor" stroke-width="2"
                 stroke-linecap="round" stroke-linejoin="round">
                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                <polyline points="14 2 14 8 20 8"/>
            </svg>
            <div class="read-share-info">
                <div class="read-share-name" title="${escHtml(displayText)}">${escHtml(displayText)}</div>
                <div class="read-share-url" title="${escHtml(displayUrl)}">${escHtml(displayUrl)}</div>
            </div>
        </div>
        <button class="btn-start-chat" id="btn-share-chat">
            <svg width="12" height="12" viewBox="0 0 24 24" fill="currentColor">
                <polygon points="5 3 19 12 5 21 5 3"/>
            </svg>
            Start Chat
        </button>`;

    section.classList.remove("hidden");

    document.getElementById("btn-share-chat").onclick = async () => {
        const btn = document.getElementById("btn-share-chat");
        btn.disabled = true;
        try {
            await onStartChat(shareInfo);
        } catch (_) {
            btn.disabled = false;
        }
    };
}

export function insertShareLinkIntoBody(link, filename) {
    return new Promise((resolve) => {
        const html = `<p style="font-family:sans-serif;margin:8px 0;">`
                   + `<a href="${link}" target="_blank" `
                   + `style="color:#0D47A1;font-weight:600;text-decoration:none;">`
                   + `📄 ${filename} — View on SmartBlue</a></p>`;

        Office.context.mailbox.item.body.setSelectedDataAsync(
            html,
            { coercionType: Office.CoercionType.Html },
            (result) => {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    Office.context.mailbox.item.body.setSelectedDataAsync(
                        `\n${filename}: ${link}\n`,
                        { coercionType: Office.CoercionType.Text },
                        () => resolve()
                    );
                } else {
                    resolve();
                }
            }
        );
    });
}

export function renderComposeResult(link) {
    document.getElementById("result-link-text").textContent = link;
    document.getElementById("compose-result").classList.remove("hidden");
    document.getElementById("compose-result").scrollIntoView({ behavior: "smooth" });
}
