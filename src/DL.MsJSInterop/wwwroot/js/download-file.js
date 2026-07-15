// Generic browser file-download helpers (ESM module).
// Loaded lazily by FileDownload__Interop via ESM import with cache-busting.

// Triggers a browser file download from a byte array passed from Blazor.
export function downloadFileFromBytes(fileName, contentType, bytes) {
    triggerDownload(new Blob([new Uint8Array(bytes)], { type: contentType }), fileName);
}

// Triggers a browser file download from a text string passed from Blazor.
export function downloadText(fileName, contentType, text) {
    triggerDownload(new Blob([text], { type: contentType }), fileName);
}

function triggerDownload(blob, fileName) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 5000);
}
