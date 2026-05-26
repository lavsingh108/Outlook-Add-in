export function showReadStatus(msg) {
    const el = document.getElementById("status-msg");
    if (el) el.innerText = msg;
}

export function showReadInitError(msg) {
    document.querySelector(".read-spinner-wrap").style.display = "none";
    document.getElementById("read-init-status").classList.add("hidden");
    document.getElementById("read-error-msg").textContent = msg;
    document.getElementById("read-init-error").classList.remove("hidden");
}

export function showComposeStatus(msg) {
    document.getElementById("compose-status").innerText = msg;
}
