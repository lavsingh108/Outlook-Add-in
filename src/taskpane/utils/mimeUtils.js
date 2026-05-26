const MIME_MAP = {
    pdf:  "application/pdf",
    doc:  "application/msword",
    docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    xls:  "application/vnd.ms-excel",
    xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    csv:  "text/csv",
    ppt:  "application/vnd.ms-powerpoint",
    pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    txt:  "text/plain",
    rtf:  "application/rtf",
    png:  "image/png",
    jpg:  "image/jpeg",
    jpeg: "image/jpeg",
    gif:  "image/gif",
    webp: "image/webp",
    zip:  "application/zip",
};

export function getMimeType(filename) {
    const ext = (filename || "").split(".").pop().toLowerCase();
    return MIME_MAP[ext] || "application/octet-stream";
}
