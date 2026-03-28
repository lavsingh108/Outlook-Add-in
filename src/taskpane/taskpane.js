/* global Office, document, fetch, FormData, atob, Uint8Array, Blob */

const AUTH_URL       = "https://ws.demo.smartblue.ai/v1/authenticate";
const UPLOAD_URL     = "https://ws.demo.smartblue.ai/v1/document/upload";
const BUNDLE_ADD_URL = "https://ws.demo.smartblue.ai/v1/document/bundle/add";
const ASK_URL        = "https://ws.demo.smartblue.ai/v1/conversation/ask/question";

let currentConversationId = null;

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        init();
    }
});

function init() {
    loadAttachments();
    document.getElementById("btn-upload-bundle").onclick = handleBundleUpload;
    document.getElementById("btn-send").onclick = sendChatMessage;
    document.getElementById("user-input").addEventListener("keydown", function (e) {
        if (e.key === "Enter" && !e.shiftKey) {
            e.preventDefault();
            sendChatMessage();
        }
    });
}

// ── Load attachments ──────────────────────────────────────────────
function loadAttachments() {
    const item        = Office.context.mailbox.item;
    const attachments = item.attachments;
    const listDiv     = document.getElementById("attachment-list");

    if (!attachments || attachments.length === 0) {
        listDiv.innerHTML = "<p style='color:#888;font-size:13px;'>No attachments found in this email.</p>";
        document.getElementById("btn-upload-bundle").disabled = true;
        return;
    }

    listDiv.innerHTML = "";
    attachments.forEach((att, index) => {
        const div = document.createElement("div");
        div.className = "att-item";
        div.innerHTML = `
            <label style="display:flex;align-items:center;gap:8px;cursor:pointer;">
                <input type="radio" name="primaryIndex" value="${index}" ${index === 0 ? "checked" : ""}/>
                <span class="att-name">${att.name}</span>
                <span class="att-size">(${formatBytes(att.size)})</span>
            </label>`;
        listDiv.appendChild(div);
    });
}

// ── Auth: Try SSO first, fall back to Exchange Identity Token ─────
async function getAuthToken() {

    // ── Method 1: Office SSO (requires Azure App Registration) ───
    try {
        const msIdToken = await Office.auth.getAccessToken({
            allowSignInPrompt:  true,
            allowConsentPrompt: true,
            forMSGraphAccess:   false
        });
        console.log("SSO token acquired, exchanging with SmartBlue...");

        const authResp = await fetch(AUTH_URL, {
            method:  "GET",
            headers: { "Authorization": "Microsoft " + msIdToken }
        });

        if (authResp.ok) {
            const authData = await authResp.json();
            if (authData.token) {
                console.log("SmartBlue session token acquired via SSO");
                return authData.token;
            }
        }
        throw new Error("SSO exchange returned no token");

    } catch (ssoErr) {
        console.warn("SSO failed (code " + ssoErr.code + "), trying Exchange identity token...");
    }

    // ── Method 2: Exchange Identity Token (works without Azure setup) ─
    return new Promise((resolve, reject) => {
        Office.context.mailbox.getUserIdentityTokenAsync(async (result) => {
            if (result.status !== Office.AsyncResultStatus.Succeeded) {
                reject(new Error("Identity token failed: " + result.error.message));
                return;
            }

            const exchangeToken = result.value;
            console.log("Exchange identity token acquired, exchanging with SmartBlue...");

            try {
                const authResp = await fetch(AUTH_URL, {
                    method:  "GET",
                    headers: { "Authorization": "Exchange " + exchangeToken }
                });

                if (!authResp.ok) {
                    const errText = await authResp.text();
                    reject(new Error("Auth failed (" + authResp.status + "): " + errText));
                    return;
                }

                const authData = await authResp.json();
                if (!authData.token) {
                    reject(new Error("No token returned from /v1/authenticate"));
                    return;
                }

                console.log("SmartBlue session token acquired via Exchange token");
                resolve(authData.token);

            } catch (fetchErr) {
                reject(new Error("Auth request failed: " + fetchErr.message));
            }
        });
    });
}

// ── Upload attachments and switch to chat ─────────────────────────
async function handleBundleUpload() {
    const item        = Office.context.mailbox.item;
    const selected    = document.querySelector("input[name='primaryIndex']:checked");
    if (!selected) { showStatus("Please select a primary document."); return; }

    const primaryIndex = parseInt(selected.value);
    const primaryAtt   = item.attachments[primaryIndex];

    showStatus("Authenticating...");
    document.getElementById("btn-upload-bundle").disabled = true;

    try {
        const token = await getAuthToken();

        showStatus("Uploading primary document...");
        const primaryBlob = await getAttachmentBlob(primaryAtt.id);
        const formData = new FormData();
        formData.append("document", primaryBlob, primaryAtt.name);

        const response = await fetch(UPLOAD_URL, {
            method:  "POST",
            headers: { Authorization: "Bearer " + token },
            body:    formData,
        });

        if (!response.ok) throw new Error("Upload failed: HTTP " + response.status);

        const data = await response.json();
        currentConversationId = data.conversation_id;

        showStatus("Uploading supporting documents...");
        for (let i = 0; i < item.attachments.length; i++) {
            if (i === primaryIndex) continue;
            const blob = await getAttachmentBlob(item.attachments[i].id);
            const sf = new FormData();
            sf.append("document", blob, item.attachments[i].name);
            await fetch(BUNDLE_ADD_URL + "?conversation_id=" + currentConversationId, {
                method:  "POST",
                headers: { Authorization: "Bearer " + token },
                body:    sf,
            });
        }

        switchToChat();

    } catch (err) {
        showStatus("Error: " + err.message);
        document.getElementById("btn-upload-bundle").disabled = false;
    }
}

// ── Get attachment as Blob ────────────────────────────────────────
function getAttachmentBlob(attachmentId) {
    return new Promise((resolve, reject) => {
        Office.context.mailbox.item.getAttachmentContentAsync(attachmentId, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                const binary = atob(result.value.content);
                const bytes  = new Uint8Array(binary.length);
                for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
                resolve(new Blob([bytes]));
            } else {
                reject(new Error(result.error.message));
            }
        });
    });
}

// ── Chat ──────────────────────────────────────────────────────────
async function sendChatMessage() {
    const input = document.getElementById("user-input");
    const text  = input.value.trim();
    if (!text) return;

    appendMessage("user", text);
    input.value = "";
    document.getElementById("btn-send").disabled = true;

    try {
        const token = await getAuthToken();
        const resp  = await fetch(ASK_URL, {
            method:  "POST",
            headers: { "Content-Type": "application/json", Authorization: "Bearer " + token },
            body:    JSON.stringify({ conversationId: currentConversationId, text: text, isMobile: false }),
        });
        const data = await resp.json();
        appendMessage("ai", data.answer || data.response || "No response received.");
    } catch (err) {
        appendMessage("ai", "Error: " + err.message);
    } finally {
        document.getElementById("btn-send").disabled = false;
    }
}

function appendMessage(role, text) {
    const hist = document.getElementById("chat-history");
    const div  = document.createElement("div");
    div.className = role === "user" ? "msg-user" : "msg-ai";
    div.innerHTML = "<strong>" + (role === "user" ? "You" : "Blue AI") + ":</strong><br>" + text;
    hist.appendChild(div);
    hist.scrollTop = hist.scrollHeight;
}

function switchToChat() {
    document.getElementById("view-attachments").classList.add("hidden");
    document.getElementById("view-chat").classList.remove("hidden");
    showStatus("");
}

function showStatus(msg) { document.getElementById("status-msg").innerText = msg; }

function formatBytes(bytes) {
    if (!bytes)       return "";
    if (bytes < 1024) return bytes + " B";
    return (bytes / 1024).toFixed(1) + " KB";
}