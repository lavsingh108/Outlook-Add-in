const AUTH_URL = "https://ws.demo.smartblue.ai/v1/authenticate";
const UPLOAD_URL = "https://ws.demo.smartblue.ai/v1/document/upload";
const BUNDLE_ADD_URL = "https://ws.demo.smartblue.ai/v1/document/bundle/add";
const ASK_URL = "https://ws.demo.smartblue.ai/v1/conversation/ask/question";

let currentConversationId = null;
let currentDocId = null;

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        init();
    }
});

async function init() {
    loadAttachments();
    document.getElementById('btn-upload-bundle').onclick = handleBundleUpload;
    document.getElementById('btn-send').onclick = sendChatMessage;
}

// 1. Get Attachments from Outlook
function loadAttachments() {
    const item = Office.context.mailbox.item;
    const attachments = item.attachments;
    const listDiv = document.getElementById('attachment-list');

    if (attachments.length === 0) {
        listDiv.innerHTML = "No attachments found in this email.";
        return;
    }

    listDiv.innerHTML = "";
    attachments.forEach((att, index) => {
        const div = document.createElement('div');
        div.className = 'att-item';
        div.innerHTML = `
            <input type="radio" name="primaryIndex" value="${index}" ${index === 0 ? 'checked' : ''}>
            <span>${att.name} (${formatBytes(att.size)})</span>
        `;
        listDiv.appendChild(div);
    });
}

// 2. Auth (Replicating your getSmartBlueToken logic)
async function getAuthToken() {
    // In a real app, you'd use Office.auth.getAccessToken() 
    // For now, we assume your backend handles the Google Identity token check
    // If you need a hardcoded testing token, return it here.
    return "YOUR_SESSION_TOKEN_HERE"; 
}

// 3. Handle Bundle Upload
async function handleBundleUpload() {
    const item = Office.context.mailbox.item;
    const primaryIndex = document.querySelector('input[name="primaryIndex"]:checked').value;
    const primaryAtt = item.attachments[primaryIndex];

    showStatus("Uploading primary document...");

    try {
        const token = await getAuthToken();
        
        // Get primary file content
        const primaryBlob = await getFileBlob(primaryAtt.id);
        const formData = new FormData();
        formData.append("document", primaryBlob, primaryAtt.name);

        const response = await fetch(UPLOAD_URL, {
            method: 'POST',
            headers: { 'Authorization': `Bearer ${token}` },
            body: formData
        });

        const data = await response.json();
        currentConversationId = data.conversation_id;
        currentDocId = data.doc_id;

        showStatus("Uploading supporting documents...");
        // Loop through others and upload to Bundle API
        for (let i = 0; i < item.attachments.length; i++) {
            if (i == primaryIndex) continue;
            const supportingBlob = await getFileBlob(item.attachments[i].id);
            const sFormData = new FormData();
            sFormData.append("document", supportingBlob, item.attachments[i].name);
            
            await fetch(`${BUNDLE_ADD_URL}?conversation_id=${currentConversationId}`, {
                method: 'POST',
                headers: { 'Authorization': `Bearer ${token}` },
                body: sFormData
            });
        }

        switchToChat();
    } catch (err) {
        showStatus("Error: " + err.message);
    }
}

// Helper: Get file from Outlook as Blob
function getFileBlob(attachmentId) {
    return new Promise((resolve, reject) => {
        Office.context.mailbox.item.getAttachmentContentAsync(attachmentId, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                // Outlook returns base64 for file attachments
                const base64 = result.value.content;
                const byteCharacters = atob(base64);
                const byteNumbers = new Array(byteCharacters.length);
                for (let i = 0; i < byteCharacters.length; i++) {
                    byteNumbers[i] = byteCharacters.charCodeAt(i);
                }
                const byteArray = new Uint8Array(byteNumbers);
                resolve(new Blob([byteArray]));
            } else {
                reject(result.error);
            }
        });
    });
}

// 4. Chat Logic
async function sendChatMessage() {
    const text = document.getElementById('user-input').value;
    if (!text) return;

    appendMsg("user", text);
    document.getElementById('user-input').value = "";

    try {
        const token = await getAuthToken();
        const resp = await fetch(ASK_URL, {
            method: 'POST',
            headers: { 
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${token}` 
            },
            body: JSON.stringify({
                conversationId: currentConversationId,
                text: text,
                isMobile: false
            })
        });

        const data = await resp.json();
        appendMsg("ai", data.answer || data.response);
    } catch (err) {
        appendMsg("ai", "Error: " + err.message);
    }
}

function appendMsg(role, text) {
    const hist = document.getElementById('chat-history');
    const div = document.createElement('div');
    div.className = role === "user" ? "msg-user" : "msg-ai";
    div.innerHTML = `<strong>${role === "user" ? "You" : "BlueAI"}:</strong><br>${text}`;
    hist.appendChild(div);
    hist.scrollTop = hist.scrollHeight;
}

function switchToChat() {
    document.getElementById('view-attachments').classList.add('hidden');
    document.getElementById('view-chat').classList.remove('hidden');
    showStatus("");
}

function showStatus(msg) {
    document.getElementById('status-msg').innerText = msg;
}

function formatBytes(bytes) {
    if (bytes < 1024) return bytes + " B";
    return (bytes / 1024).toFixed(1) + " KB";
}