const PROP_KEY = "conversationsMap";

export function loadCustomProps() {
    return new Promise((resolve, reject) => {
        Office.context.mailbox.item.loadCustomPropertiesAsync(result => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                resolve(result.value);
            } else {
                reject(new Error(result.error?.message || "Failed to load custom properties"));
            }
        });
    });
}

export function getConversationMap(customProps) {
    try {
        const raw = customProps.get(PROP_KEY);
        return raw ? JSON.parse(raw) : {};
    } catch {
        return {};
    }
}

export function saveConversationRecord(customProps, fingerprint, record) {
    const map = getConversationMap(customProps);
    map[fingerprint] = record;
    customProps.set(PROP_KEY, JSON.stringify(map));
    return new Promise((resolve, reject) => {
        customProps.saveAsync(result => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                resolve();
            } else {
                reject(new Error(result.error?.message || "Failed to save custom properties"));
            }
        });
    });
}

export function singleFingerprint(att) {
    return `att_${att.id}`;
}

export function bundleFingerprint(primaryAtt, secondaryAtts) {
    const secondaryIds = secondaryAtts.map(a => a.id).sort();
    return `bundle_${[primaryAtt.id, ...secondaryIds].join("_")}`;
}
