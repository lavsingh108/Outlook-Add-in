// ── Proxy base URL ────────────────────────────────────────────────
// Change this to wherever you deploy the proxy server.
const PROXY_BASE     = "https://headphone-crust-stipulate.ngrok-free.dev";

const AUTH_URL       = `${PROXY_BASE}/v1/authenticate`;
const UPLOAD_URL     = `${PROXY_BASE}/v1/document/upload`;
const BUNDLE_ADD_URL = `${PROXY_BASE}/v1/document/bundle/add`;
const WELCOME_URL    = `${PROXY_BASE}/v1/conversation/ask/welcome`;
const ASK_URL        = `${PROXY_BASE}/v1/conversation/ask/question`;

// ── MSAL Config ───────────────────────────────────────────────────
const AZURE_CLIENT_ID = "c49037f2-0565-4a5c-8b17-f9b8b3ee35c7";
const AZURE_TENANT_ID = "f895e126-dbc8-41bb-b00b-5cd2172346f9";
const SCOPES = ["openid", "profile", "email", "User.Read"];

const msalConfig = {
    auth: {
        clientId: AZURE_CLIENT_ID,
        authority: "https://login.microsoftonline.com/" + AZURE_TENANT_ID,
        redirectUri: window.location.href.split("?")[0]
    },
    cache: {
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: false
    }
};

let _msal = null;
function getMsal() {
    if (!_msal) _msal = new msal.PublicClientApplication(msalConfig);
    return _msal;
}

// Cache the SmartBlue session token for the lifetime of the taskpane session.
// Avoids re-authenticating on every upload/chat call.
let _cachedSmartBlueToken = null;
let currentConversationId = null;

// ── Entry Point - Office Ready ────────────────────────────────────
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        init();
    }
});

function init() {
    loadAttachments();
    document.getElementById("btn-upload-bundle").onclick = handleBundleUpload;
    document.getElementById("btn-send").onclick = sendChatMessage;
    document.getElementById("btn-back").onclick = switchToAttachments;
    document.getElementById("user-input").addEventListener("keydown", function (e) {
        if (e.key === "Enter" && !e.shiftKey) {
            e.preventDefault();
            sendChatMessage();
        }
    });
}

// ── Load attachments ──────────────────────────────────────────────
function loadAttachments() {
    const item = Office.context.mailbox.item;
    const attachments = item.attachments;
    const listDiv = document.getElementById("attachment-list");

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

// ── Auth: MSAL popup → proxy → SmartBlue session token ───────────
async function getAuthToken() {
    // Return cached token if we already have one for this session
    if (_cachedSmartBlueToken) return _cachedSmartBlueToken;

    const msalInstance = getMsal();
    let idToken = null;

    // 1. Try silent first (no popup if session is cached)
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
        try {
            const silent = await msalInstance.acquireTokenSilent({
                scopes:  SCOPES,
                account: accounts[0]
            });
            idToken = silent.idToken;
            console.log("MSAL silent token acquired for:", accounts[0].username);
        } catch (silentErr) {
            console.warn("Silent failed, falling back to popup:", silentErr.message);
        }
    }

    // 2. Popup login if no cached session
    if (!idToken) {
        try {
            const popup = await msalInstance.loginPopup({
                scopes: SCOPES,
                prompt: "select_account"
            });
            idToken = popup.idToken;
            console.log("MSAL popup login OK:", popup.account.username);
        } catch (popupErr) {
            console.error("MSAL popup error:", popupErr);
            throw new Error("Sign-in failed: " + (popupErr.message || popupErr.errorCode));
        }
    }

    // 3. Exchange Microsoft ID token for SmartBlue session token via proxy
    console.log("Exchanging Microsoft ID token via proxy...");
    const authResp = await fetch(AUTH_URL, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ idToken }),
    });

    if (!authResp.ok) {
        const err = await authResp.text();
        throw new Error("Auth exchange failed (" + authResp.status + "): " + err);
    }

    const authData = await authResp.json();

    if (!authData.token) {
        throw new Error("Proxy returned no token. Response: " + JSON.stringify(authData));
    }

    console.log("SmartBlue session token acquired via proxy.");
    _cachedSmartBlueToken = authData.token;
    return _cachedSmartBlueToken;
}

// ── Upload pipeline ───────────────────────────────────────────────
async function handleBundleUpload() {
    const item = Office.context.mailbox.item;
    const selected = document.querySelector("input[name='primaryIndex']:checked");
    if (!selected) { showStatus("Please select a primary document."); return; }

    const primaryIndex = parseInt(selected.value);
    const primaryAtt   = item.attachments[primaryIndex];

    showStatus("Signing in...");
    document.getElementById("btn-upload-bundle").disabled = true;

    try {
        const token = await getAuthToken();

        // ── Upload primary document ───────────────────────────────
        showStatus("Uploading primary document...");
        const primaryBlob = await getAttachmentBlob(primaryAtt.id, primaryAtt.name);
        const formData = new FormData();
        formData.append("document", primaryBlob, primaryAtt.name);

        const uploadResp = await fetch(UPLOAD_URL, {
            method:  "POST",
            headers: { Authorization: "Bearer " + token },
            body:    formData,
            // Do NOT set Content-Type — the browser sets multipart boundary automatically
        });

        if (!uploadResp.ok) {
            const detail = await uploadResp.text();
            throw new Error("Upload failed (" + uploadResp.status + "): " + detail);
        }

        const uploadData = await uploadResp.json();
        currentConversationId = uploadData.conversation_id;
        currentDocumentId = uploadData.document_id;
        console.log("Primary document uploaded. conversation_id:", currentConversationId);

        // ── Upload supporting documents ───────────────────────────
        const supporting = item.attachments.filter((_, i) => i !== primaryIndex);
        if (supporting.length > 0) {
            showStatus(`Uploading ${supporting.length} supporting document(s)...`);
            for (const att of supporting) {
                const blob = await getAttachmentBlob(att.id, att.name);
                const sf = new FormData();
                sf.append("document", blob, att.name);
                const bundleResp = await fetch(
                    `${BUNDLE_ADD_URL}?conversation_id=${encodeURIComponent(currentConversationId)}`,
                    {
                        method:  "POST",
                        headers: { Authorization: "Bearer " + token },
                        body:    sf,
                    }
                );
                if (!bundleResp.ok) {
                    console.warn("Bundle add failed for:", att.name, await bundleResp.text());
                }
            }
        }

        const welcomeResp = await fetch(WELCOME_URL, {
            conversationId: currentConversationId,
            documentId: currentDocumentId
        },{
            method:  "POST",
            headers: { Authorization: "Bearer " + token }
        });
        const welcomeData = await welcomeResp.json();
        appendMessage("ai", welcomeData.answer || welcomeData.response || "No response received.");

        switchToChat();

    } catch (err) {
        console.error("Upload error:", err);
        showStatus("Error: " + err.message);
        // Clear cached token in case it expired mid-session
        _cachedSmartBlueToken = null;
        document.getElementById("btn-upload-bundle").disabled = false;
    }
}

// ── MIME type lookup by file extension ───────────────────────────
// Outlook's getAttachmentContentAsync returns raw bytes with no type
// metadata, so we derive the MIME type from the filename ourselves.
function getMimeType(filename) {
    const ext = (filename || "").split(".").pop().toLowerCase();
    const MIME_MAP = {
        // Documents
        pdf:  "application/pdf",
        doc:  "application/msword",
        docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        // Spreadsheets
        xls:  "application/vnd.ms-excel",
        xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        csv:  "text/csv",
        // Presentations
        ppt:  "application/vnd.ms-powerpoint",
        pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        // Text
        txt:  "text/plain",
        rtf:  "application/rtf",
        // Images
        png:  "image/png",
        jpg:  "image/jpeg",
        jpeg: "image/jpeg",
        gif:  "image/gif",
        webp: "image/webp",
        // Archives
        zip:  "application/zip",
        // Fallback — SmartBlue should accept this for anything not listed
        "":   "application/octet-stream",
    };
    return MIME_MAP[ext] || "application/octet-stream";
}

// ── Get attachment as Blob ────────────────────────────────────────
// filename is required so we can set the correct MIME type on the Blob.
function getAttachmentBlob(attachmentId, filename) {
    return new Promise((resolve, reject) => {
        Office.context.mailbox.item.getAttachmentContentAsync(attachmentId, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                const binary  = atob(result.value.content);
                const bytes   = new Uint8Array(binary.length);
                for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
                const mimeType = getMimeType(filename);
                console.log(`Attachment "${filename}" → MIME type: ${mimeType}`);
                resolve(new Blob([bytes], { type: mimeType }));
            } else {
                reject(new Error(result.error.message));
            }
        });
    });
}

// ── Chat ──────────────────────────────────────────────────────────
async function sendChatMessage() {
    const input = document.getElementById("user-input");
    const text = input.value.trim();
    if (!text) return;

    appendMessage("user", text);
    input.value = "";
    document.getElementById("btn-send").disabled = true;

    try {
        const token = await getAuthToken();
        const resp = await fetch(ASK_URL, {
            method:  "POST",
            headers: {
                "Content-Type": "application/json",
                Authorization:  "Bearer " + token,
            },
            body: JSON.stringify({
                conversationId: currentConversationId,
                text,
                isMobile: false,
            }),
        });

        if (!resp.ok) {
            const detail = await resp.text();
            throw new Error("Ask failed (" + resp.status + "): " + detail);
        }

        const data = await resp.json();
        appendMessage("ai", data.answer || data.response || "No response received.");
    } catch (err) {
        console.error("Chat error:", err);
        appendMessage("ai", "Error: " + err.message);
        // Clear cached token in case it expired
        _cachedSmartBlueToken = null;
    } finally {
        document.getElementById("btn-send").disabled = false;
    }
}

// ── Response formatter ────────────────────────────────────────────
/**
 * Converts the SmartBlue API response string to clean HTML.
 *
 * Handles:
 *  1. <blueEmbed-doc-page>UUID:file.pdf:N</blueEmbed-doc-page>
 *     → <span class="page-ref">pg N</span>
 *  2. **bold** → <strong>bold</strong>
 *  3. Lines starting with "* " or "● " → <ul><li> list items
 *  4. Remaining non-empty lines → <p> paragraphs
 */
function formatResponse(raw) {
    // 1. Replace citation tags — extract only the page number
    let text = raw.replace(
        /<blueEmbed-doc-page>[^:]+:[^:]+:(\d+)<\/blueEmbed-doc-page>/g,
        '<span class="page-ref">pg $1</span>'
    );

    // 2. Bold
    text = text.replace(/\*\*(.+?)\*\*/g, "<strong>$1</strong>");

    // 3. Walk lines and build HTML
    const lines = text.split(/\n/);
    let html = "";
    let inList = false;

    for (const raw of lines) {
        const line = raw.trim();
        if (!line) {
            // Blank line closes an open list
            if (inList) { html += "</ul>"; inList = false; }
            continue;
        }
        if (/^[*●•]\s+/.test(line)) {
            if (!inList) { html += '<ul class="ai-list">'; inList = true; }
            html += "<li>" + line.replace(/^[*●•]\s+/, "") + "</li>";
        } else {
            if (inList) { html += "</ul>"; inList = false; }
            html += "<p>" + line + "</p>";
        }
    }
    if (inList) html += "</ul>";

    return html;
}

// ── UI helpers ────────────────────────────────────────────────────
function appendMessage(role, text) {
    const hist = document.getElementById("chat-history");
    const div  = document.createElement("div");

    if (role === "user") {
        div.className = "msg-user";
        // User messages: plain text, escape HTML
        const p = document.createElement("p");
        p.textContent = text;
        div.appendChild(p);
    } else {
        div.className = "msg-ai";
        // AI messages: full markdown + citation rendering
        div.innerHTML = formatResponse(text);
    }

    hist.appendChild(div);
    // Scroll to the new message
    hist.scrollTop = hist.scrollHeight;
}

function switchToChat() {
    document.getElementById("view-attachments").classList.add("hidden");
    document.getElementById("view-chat").classList.remove("hidden");
    document.getElementById("btn-back").classList.remove("hidden");
    showStatus("");
}

function switchToAttachments() {
    document.getElementById("view-chat").classList.add("hidden");
    document.getElementById("view-attachments").classList.remove("hidden");
    document.getElementById("btn-back").classList.add("hidden");
    // Re-enable upload button and clear status in case of prior error
    document.getElementById("btn-upload-bundle").disabled = false;
    showStatus("");
}

function showStatus(msg) { document.getElementById("status-msg").innerText = msg; }

function formatBytes(bytes) {
    if (!bytes) return "";
    if (bytes < 1024) return bytes + " B";
    return (bytes / 1024).toFixed(1) + " KB";
}