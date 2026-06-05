// ── Config & Constants ──────────────────────────────────────────────────────

export const PROXY_BASE = "https://headphone-crust-stipulate.ngrok-free.dev"; // This will be the endpoint of the BlueAI backend (this Proxy will forward requests to the BlueAI backend, while also handling CORS and authentication)
export const BLUE_BASE  = "https://demo.smartblue.ai";

// Domains whose /conversation URLs are treated as SmartBlue share links.
// Only emails containing a link on one of these domains will show "Start Chat".
export const SMARTBLUE_DOMAINS = ["demo.smartblue.ai", "try.smartblue.ai", "app.smartblue.ai"];

export const AUTH_URL         = `${PROXY_BASE}/v1/authenticate`;
export const UPLOAD_URL       = `${PROXY_BASE}/v1/document/upload`;
export const BUNDLE_ADD_URL   = `${PROXY_BASE}/v1/document/bundle/add`;
export const SHARE_URL        = `${PROXY_BASE}/v1/document/share`;
export const WELCOME_URL      = `${PROXY_BASE}/v1/conversation/ask/welcome`;
export const ASK_URL          = `${PROXY_BASE}/v1/conversation/ask/question`;
export const CONVERSATION_URL = `${PROXY_BASE}/v1/conversation`;

export const AZURE_CLIENT_ID = "c49037f2-0565-4a5c-8b17-f9b8b3ee35c7";
export const AZURE_TENANT_ID = "f895e126-dbc8-41bb-b00b-5cd2172346f9";
export const SCOPES          = ["openid", "profile", "email", "User.Read"];

export const msalConfig = {
    auth: {
        clientId:    AZURE_CLIENT_ID,
        authority:   "https://login.microsoftonline.com/common", // multi-tenant
        redirectUri: window.location.href.split("?")[0],
    },
    cache: { cacheLocation: "sessionStorage", storeAuthStateInCookie: false },
};
