// IndexedDB-backed cache for country flag images (PNG bytes keyed by ISO2 code).
// Used by browser hosts (Blazor WASM, web add-ins, interactive server circuits) so
// flags are not re-downloaded and remain available offline.

const DB_NAME = 'etw-flag-images';
const DB_VERSION = 1;
const STORE_NAME = 'flags';

let dbPromise = null;

function openDb() {
    if (!dbPromise) {
        dbPromise = new Promise((resolve, reject) => {
            const request = indexedDB.open(DB_NAME, DB_VERSION);
            request.onupgradeneeded = () => {
                if (!request.result.objectStoreNames.contains(STORE_NAME)) {
                    request.result.createObjectStore(STORE_NAME);
                }
            };
            request.onsuccess = () => resolve(request.result);
            request.onerror = () => reject(request.error);
        });
    }
    return dbPromise;
}

function runTransaction(mode, action) {
    return openDb().then(db => new Promise((resolve, reject) => {
        const tx = db.transaction(STORE_NAME, mode);
        const store = tx.objectStore(STORE_NAME);
        const request = action(store);
        tx.oncomplete = () => resolve(request ? request.result : undefined);
        tx.onerror = () => reject(tx.error);
        tx.onabort = () => reject(tx.error);
    }));
}

export async function getImage(iso2) {
    const value = await runTransaction('readonly', store => store.get(iso2));
    if (!value) return null;
    return value instanceof Uint8Array ? value : new Uint8Array(value);
}

export function saveImage(iso2, bytes) {
    return runTransaction('readwrite', store => store.put(bytes, iso2));
}

export function clearImages() {
    return runTransaction('readwrite', store => store.clear());
}
