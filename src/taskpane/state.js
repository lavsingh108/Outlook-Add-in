// ── Application State ───────────────────────────────────────────────────────
// All mutable runtime state lives here so it can be imported by any module
// without creating circular dependencies.

export const state = {
    currentConversationId:    null,
    currentDocumentId:        null,
    suppressAttachmentRefresh: false,
};

// MSAL / auth
export let _msal                 = null;
export let _cachedSmartBlueToken = null;

// Custom properties cache
export let _customProps = null;

// Compose-mode state
export let _composeAttachments        = [];
export let _composeRecipients         = [];
export let _senderEmail               = "";
export let _composeConversationCtx    = null;       // { conversationId, documentId } after first upload
export let _composeUploadedAttIds     = new Set();  // att.id values uploaded this compose session
export let _composeSharedRecipients   = new Set();  // recipient emails already shared with
export let _composeRefreshTimer       = null;       // debounce timer for change events
export let _composeAccessLevel        = "restricted"; // "restricted" | "anonymous"

// Read-mode state
export let _readShareInfo   = null;  // share link found in email body; set in initRead
export let _chatFromCompose = false; // true when chat opened from compose mode
export let _chatFromSent    = false; // true when chat opened from sent mode

// ── Setters (keeps mutation explicit and grep-friendly) ────────────────────
export function setMsal(instance)               { _msal = instance; }
export function setCachedToken(token)           { _cachedSmartBlueToken = token; }
export function setCustomProps(cp)              { _customProps = cp; }
export function setComposeAttachments(list)     { _composeAttachments = list; }
export function setComposeRecipients(list)      { _composeRecipients = list; }
export function setSenderEmail(email)           { _senderEmail = email; }
export function setComposeConversationCtx(ctx)  { _composeConversationCtx = ctx; }
export function setComposeUploadedAttIds(set)   { _composeUploadedAttIds = set; }
export function setComposeSharedRecipients(set) { _composeSharedRecipients = set; }
export function setComposeRefreshTimer(id)      { _composeRefreshTimer = id; }
export function setComposeAccessLevel(level)    { _composeAccessLevel = level; }
export function setReadShareInfo(info)          { _readShareInfo = info; }
export function setChatFromCompose(val)         { _chatFromCompose = val; }
export function setChatFromSent(val)            { _chatFromSent = val; }
