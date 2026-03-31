const AUTH_URL       = "https://ws.demo.smartblue.ai/v1/authenticate";
const UPLOAD_URL     = "https://ws.demo.smartblue.ai/v1/document/upload";
const BUNDLE_ADD_URL = "https://ws.demo.smartblue.ai/v1/document/bundle/add";
const ASK_URL        = "https://ws.demo.smartblue.ai/v1/conversation/ask/question";

// ── MSAL Config ───────────────────────────────────────────────────
const AZURE_CLIENT_ID = "c49037f2-0565-4a5c-8b17-f9b8b3ee35c7";  // Azure Portal → App registrations → Application (client) ID
const AZURE_TENANT_ID = "f895e126-dbc8-41bb-b00b-5cd2172346f9";  // Azure Portal → App registrations → Directory (tenant) ID
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

// ── Auth: MSAL.js popup → SmartBlue session token ────────────────
async function getAuthToken() {
    const msalInstance = getMsal();

    let idToken = null;
    let accessToken = null;

    // 1. Try silent (cached session — no popup shown)
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
        try {
            const silent = await msalInstance.acquireTokenSilent({
                scopes:  SCOPES,
                account: accounts[0]
            });
            idToken = silent.idToken;
            accessToken = silent.accessToken;
            console.log("MSAL silent token OK:", accounts[0].username);
        } catch (silentErr) {
            console.warn("Silent failed, will try popup:", silentErr.message);
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

    // 3. Exchange Microsoft ID token for SmartBlue session token
    console.log("Exchanging Microsoft ID token with SmartBlue...");

    const authResp = await fetch(AUTH_URL, {
        method: "GET",
        headers: { "Authorization": "Microsoft " + idToken }
    });

    console.log("Auth response status:", authResp.status);
    const rawText = await authResp.text();
    console.log("Auth response body:", rawText);

    if (!authResp.ok) {
        throw new Error("Auth exchange failed (" + authResp.status + "): " + rawText);
    }

    let authData;
    try { authData = JSON.parse(rawText); }
    catch (e) { throw new Error("Auth response not JSON: " + rawText); }

    const sessionToken = authData.token
        || authData.access_token
        || authData.accessToken
        || authData.sessionToken;

    if (!sessionToken) {
        throw new Error("No token in auth response. Got keys: " + Object.keys(authData).join(", "));
    }

    console.log("SmartBlue session token acquired");
    return sessionToken;
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

        showStatus("Uploading primary document...");
        const primaryBlob = await getAttachmentBlob(primaryAtt.id);
        const formData = new FormData();
        formData.append("document", primaryBlob, primaryAtt.name);

        const response = await fetch(UPLOAD_URL, {
            method: "POST",
            headers: { Authorization: "Bearer " + token },
            body: formData,
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
                method: "POST",
                headers: { Authorization: "Bearer " + token },
                body: sf,
            });
        }

        switchToChat();

    } catch (err) {
        console.error("Upload error:", err);
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
    const text = input.value.trim();
    if (!text) return;

    appendMessage("user", text);
    input.value = "";
    document.getElementById("btn-send").disabled = true;

    try {
        const token = await getAuthToken();
        const resp = await fetch(ASK_URL, {
            method: "POST",
            headers: { "Content-Type": "application/json", Authorization: "Bearer " + token },
            body: JSON.stringify({ conversationId: currentConversationId, text: text, isMobile: false }),
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
    const div = document.createElement("div");
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
    if (!bytes) return "";
    if (bytes < 1024) return bytes + " B";
    return (bytes / 1024).toFixed(1) + " KB";
}