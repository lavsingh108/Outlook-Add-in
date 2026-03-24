const API_BASE = "https://ws.demo.smartblue.ai/v1";
let currentConversationId = null;
let currentDocId = null;

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        loadAttachments();
        document.getElementById('btn-upload-bundle').onclick = uploadBundle;
        document.getElementById('btn-send').onclick = onSendChat;
    }
});

async function loadAttachments() {
    const item = Office.context.mailbox.item;
    const attachments = item.attachments;
    const listDiv = document.getElementById('attachment-list');
    
    if (attachments.length === 0) {
        listDiv.innerHTML = "No attachments found.";
        return;
    }

    listDiv.innerHTML = "";
    attachments.forEach((att, index) => {
        const row = document.createElement('div');
        row.className = 'attachment-item';
        row.innerHTML = `
            <span>${att.name}</span>
            <input type="radio" name="primary" value="${index}" ${index === 0 ? 'checked' : ''}>
        `;
        listDiv.appendChild(row);
    });
}

// Outlook requires a callback to get attachment content
async function getAttachmentBlob(attachment) {
    return new Promise((resolve, reject) => {
        Office.context.mailbox.item.getAttachmentContentAsync(attachment.id, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                // Convert content to blob
                const content = result.value.content;
                const blob = new Blob([content], { type: attachment.contentType });
                resolve(blob);
            } else {
                reject(result.error);
            }
        });
    });
}

async function uploadBundle() {
    const item = Office.context.mailbox.item;
    const selectedIndex = document.querySelector('input[name="primary"]:checked').value;
    const attachments = item.attachments;

    try {
        // 1. Get Auth Token (You'll need to adapt your specific Auth flow here)
        const authToken = "YOUR_TOKEN_LOGIC"; 

        // 2. Upload Primary
        const primaryBlob = await getAttachmentBlob(attachments[selectedIndex]);
        const formData = new FormData();
        formData.append("document", primaryBlob, attachments[selectedIndex].name);

        const resp = await fetch(`${API_BASE}/document/upload`, {
            method: 'POST',
            headers: { 'Authorization': `Bearer ${authToken}` },
            body: formData
        });
        const primaryData = await resp.json();
        
        currentConversationId = primaryData.conversation_id;
        currentDocId = primaryData.doc_id;

        // 3. Switch to Chat View
        document.getElementById('view-attachments').classList.add('hidden');
        document.getElementById('view-chat').classList.remove('hidden');
        addChatMessage("BlueAI", "Document uploaded. How can I help?");

    } catch (err) {
        alert("Upload failed: " + err.message);
    }
}

async function onSendChat() {
    const text = document.getElementById('chat-input').value;
    if (!text) return;

    addChatMessage("User", text);
    document.getElementById('chat-input').value = "";

    const resp = await fetch(`${API_BASE}/conversation/ask/question`, {
        method: 'POST',
        headers: { 
            'Content-Type': 'application/json',
            'Authorization': `Bearer YOUR_TOKEN` 
        },
        body: JSON.stringify({
            conversationId: currentConversationId,
            text: text
        })
    });
    
    const data = await resp.json();
    addChatMessage("BlueAI", data.answer || data.response);
}

function addChatMessage(sender, text) {
    const history = document.getElementById('chat-history');
    const msg = document.createElement('div');
    msg.className = sender === "User" ? "user-msg" : "ai-msg";
    msg.innerHTML = `<strong>${sender}:</strong> ${text}`;
    history.appendChild(msg);
    history.scrollTop = history.scrollHeight;
}