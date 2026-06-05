// ── Storage Service ─────────────────────────────────────────────────────────
// Wrappers around Office custom properties (per-item) and roamingSettings
// (cross-device, thread-scoped).

import { setComposeAccessLevel } from "../state.js";

const PROP_KEY = "conversationsMap";

// ── Custom Properties (per mail item) ─────────────────────────────────────

export function loadCustomProps() {
    return new Promise((resolve, reject) => {
        Office.context.mailbox.item.loadCustomPropertiesAsync(result => {
            if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
            else reject(new Error(result.error?.message || "Failed to load custom properties"));
        });
    });
}

export function getConversationMap(cp) {
    try {
        const raw = cp.get(PROP_KEY);
        return raw ? JSON.parse(raw) : {};
    } catch {
        return {};
    }
}

export function saveConversationRecord(cp, fingerprint, record) {
    const map = getConversationMap(cp);
    map[fingerprint] = record;
    cp.set(PROP_KEY, JSON.stringify(map));
    return new Promise((resolve, reject) => {
        cp.saveAsync(result => {
            if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
            else reject(new Error(result.error?.message || "Failed to save custom properties"));
        });
    });
}

// ── Fingerprint helpers ────────────────────────────────────────────────────

export function singleFingerprint(att) {
    return `att_${att.id}`;
}

export function bundleFingerprint(primaryAtt, secondaryAtts) {
    const ids = secondaryAtts.map(a => a.id).sort();
    return `bundle_${[primaryAtt.id, ...ids].join("_")}`;
}

// ── Attachment record helpers ──────────────────────────────────────────────
// These need customProps injected rather than reading from global state so
// they remain pure and testable.

export function getAttachmentRecord(att, customProps) {
    if (!customProps) return null;
    const map = getConversationMap(customProps);
    if (map[singleFingerprint(att)]) return map[singleFingerprint(att)];
    const bundleKey = Object.keys(map).find(
        fp => fp.startsWith("bundle_") && fp.includes(att.id)
    );
    return bundleKey ? map[bundleKey] : null;
}

export function isAttachmentUploaded(att, customProps) {
    return !!getAttachmentRecord(att, customProps);
}

// ── RoamingSettings (thread-scoped, cross-device) ─────────────────────────
//
// Key:   "sb_thread_{outlookConversationId}"
// Value: Array of { conversationId, documentId, label, uploadType, timestamp }
//
// 32 KB total limit. Each record ≈ 150 bytes → ~200 threads before pruning.
// Entries older than 90 days are pruned on every write.

function getThreadKey() {
    const threadId = Office.context.mailbox.item.conversationId || "unknown";
    return `sb_thread_${threadId}`;
}

export function saveThreadContext(record) {
    try {
        const rs  = Office.context.roamingSettings;
        const key = getThreadKey();

        // Merge by conversationId — no duplicates
        const existing = getThreadContextAll();
        const map = {};
        existing.forEach(r => { map[r.conversationId] = r; });
        map[record.conversationId] = record;
        rs.set(key, Object.values(map));

        // Maintain a key index so we can prune old threads
        const cutoff    = Date.now() - 90 * 24 * 60 * 60 * 1000;
        const index     = rs.get("sb_thread_index") || [];
        if (!index.includes(key)) index.push(key);

        const activeIdx = index.filter(k => {
            const recs = rs.get(k);
            if (!recs || !recs.length) { rs.remove(k); return false; }
            const latest = Math.max(...recs.map(r => r.timestamp || 0));
            if (latest < cutoff) { rs.remove(k); return false; }
            return true;
        });
        rs.set("sb_thread_index", activeIdx);

        rs.saveAsync(result => {
            if (result.status !== Office.AsyncResultStatus.Succeeded)
                console.warn("roamingSettings save failed:", result.error?.message);
            else
                console.log("Thread context saved to roamingSettings:", key);
        });
    } catch (e) {
        console.warn("saveThreadContext failed:", e.message);
    }
}

export function getThreadContextAll() {
    try {
        return Office.context.roamingSettings.get(getThreadKey()) || [];
    } catch {
        return [];
    }
}

// ── Compose preferences (persisted via roamingSettings) ───────────────────

export function loadComposePrefs() {
    try {
        const rs            = Office.context.roamingSettings;
        const removeChecked = rs.get("sb_pref_remove_attachment");
        const accessLevel   = rs.get("sb_pref_access_level");

        if (removeChecked !== undefined && removeChecked !== null) {
            const chk = document.getElementById("chk-include-attachment");
            if (chk) chk.checked = !!removeChecked;
        }
        if (accessLevel) {
            setComposeAccessLevel(accessLevel);
            const sel = document.getElementById("sel-access");
            if (sel) sel.value = accessLevel;
        }
    } catch (e) {
        console.warn("loadComposePrefs failed:", e.message);
    }
}

export function saveComposePrefs() {
    try {
        const rs  = Office.context.roamingSettings;
        const chk = document.getElementById("chk-include-attachment");
        const sel = document.getElementById("sel-access");
        if (chk) rs.set("sb_pref_remove_attachment", chk.checked);
        if (sel) rs.set("sb_pref_access_level", sel.value);
        rs.saveAsync(r => {
            if (r.status !== Office.AsyncResultStatus.Succeeded)
                console.warn("saveComposePrefs failed:", r.error?.message);
        });
    } catch (e) {
        console.warn("saveComposePrefs failed:", e.message);
    }
}
