// ── Auth Service ────────────────────────────────────────────────────────────
// MSAL initialisation + SmartBlue token acquisition.
//
// NOTE on live bindings: ES module imports of `let` variables are live —
// _msal and _cachedSmartBlueToken always reflect the latest value from state.js
// without needing re-import tricks.

import { msalConfig, SCOPES, AUTH_URL } from "../config.js";
import { _msal, _cachedSmartBlueToken, setMsal, setCachedToken } from "../state.js";

export function clearToken() {
    setCachedToken(null);
}

/**
 * Obtain a SmartBlue bearer token.
 * 1. Returns cached token if available.
 * 2. Attempts a silent MSAL token refresh.
 * 3. Falls back to a login popup.
 * 4. Exchanges the Microsoft id_token for a SmartBlue token via the proxy.
 */
export async function getAuthToken() {
    // _cachedSmartBlueToken is a live binding — always reflects current state
    if (_cachedSmartBlueToken) return _cachedSmartBlueToken;

    // Initialise MSAL on first use
    if (!_msal) setMsal(new msal.PublicClientApplication(msalConfig));

    let idToken = null;
    const accounts = _msal.getAllAccounts();
    if (accounts.length > 0) {
        try {
            const silent = await _msal.acquireTokenSilent({ scopes: SCOPES, account: accounts[0] });
            idToken = silent.idToken;
        } catch (e) {
            console.warn("Silent token failed:", e.message);
        }
    }

    if (!idToken) {
        try {
            const popup = await _msal.loginPopup({ scopes: SCOPES, prompt: "select_account" });
            idToken = popup.idToken;
        } catch (e) {
            throw new Error("Sign-in failed: " + (e.message || e.errorCode));
        }
    }

    const authResp = await fetch(AUTH_URL, {
        method:  "POST",
        headers: { "Content-Type": "application/json" },
        body:    JSON.stringify({ idToken }),
    });
    if (!authResp.ok) throw new Error("Auth failed (" + authResp.status + "): " + await authResp.text());

    const { token } = await authResp.json();
    if (!token) throw new Error("No token returned from auth proxy.");

    setCachedToken(token);
    return token;
}
