/* global Office, document, fetch, FormData, atob, Uint8Array, Blob */

const AUTH_URL = "https://ws.demo.smartblue.ai/v1/authenticate";
const UPLOAD_URL = "https://ws.demo.smartblue.ai/v1/document/upload";
const BUNDLE_ADD_URL = "https://ws.demo.smartblue.ai/v1/document/bundle/add";
const ASK_URL = "https://ws.demo.smartblue.ai/v1/conversation/ask/question";

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

// ── Load attachments from the current email ──────────────────────
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

// ── Get auth token ────────────────────────────────────────────────
async function getAuthToken() {
    // Replace with your real auth logic / Office SSO token
    const token = await Office.auth.getAccessToken();
    console.log("Token:: ", token);
    
    return await Office.auth.getAccessToken();
    // return "YOUR_SESSION_TOKEN_HERE";
}

// ── Upload attachments and switch to chat ─────────────────────────
async function handleBundleUpload() {
    const item = Office.context.mailbox.item;
    const selected = document.querySelector("input[name='primaryIndex']:checked");
    if (!selected) {
        showStatus("Please select a primary document.");
        return;
    }
    const primaryIndex = parseInt(selected.value);
    const primaryAtt = item.attachments[primaryIndex];

    showStatus("Uploading primary document...");
    document.getElementById("btn-upload-bundle").disabled = true;

    try {
        const token = await getAuthToken();

        // Upload primary document
        const primaryBlob = await getAttachmentBlob(primaryAtt.id);
        const formData = new FormData();
        formData.append("document", primaryBlob, primaryAtt.name);

        const response = await fetch(UPLOAD_URL, {
            method: "POST",
            headers: { Authorization: `Bearer ${token}` },
            body: formData,
        });

        const data = await response.json();
        currentConversationId = data.conversation_id;

        // Upload remaining attachments as supporting docs
        showStatus("Uploading supporting documents...");
        for (let i = 0; i < item.attachments.length; i++) {
            if (i === primaryIndex) continue;
            const blob = await getAttachmentBlob(item.attachments[i].id);
            const sf = new FormData();
            sf.append("document", blob, item.attachments[i].name);
            await fetch(`${BUNDLE_ADD_URL}?conversation_id=${currentConversationId}`, {
                method: "POST",
                headers: { Authorization: `Bearer ${token}` },
                body: sf,
            });
        }

        switchToChat();
    } catch (err) {
        showStatus("Error: " + err.message);
        document.getElementById("btn-upload-bundle").disabled = false;
    }
}

// ── Get attachment content as Blob ────────────────────────────────
function getAttachmentBlob(attachmentId) {
    return new Promise((resolve, reject) => {
        Office.context.mailbox.item.getAttachmentContentAsync(attachmentId, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                const base64 = result.value.content;
                const byteChars = atob(base64);
                const byteArr = new Uint8Array(byteChars.length);
                for (let i = 0; i < byteChars.length; i++) {
                    byteArr[i] = byteChars.charCodeAt(i);
                }
                resolve(new Blob([byteArr]));
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
            headers: {
                "Content-Type": "application/json",
                Authorization: `Bearer ${token}`,
            },
            body: JSON.stringify({
                conversationId: currentConversationId,
                text: text,
                isMobile: false,
            }),
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
    div.innerHTML = `<strong>${role === "user" ? "You" : "Blue AI"}:</strong><br>${text}`;
    hist.appendChild(div);
    hist.scrollTop = hist.scrollHeight;
}

// ── UI helpers ────────────────────────────────────────────────────
function switchToChat() {
    document.getElementById("view-attachments").classList.add("hidden");
    document.getElementById("view-chat").classList.remove("hidden");
    showStatus("");
}

function showStatus(msg) {
    document.getElementById("status-msg").innerText = msg;
}

function formatBytes(bytes) {
    if (!bytes) return "";
    if (bytes < 1024) return bytes + " B";
    return (bytes / 1024).toFixed(1) + " KB";
}
