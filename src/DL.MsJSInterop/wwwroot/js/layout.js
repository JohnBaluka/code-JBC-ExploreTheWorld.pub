// Generic browser window/layout helpers (ESM module).
// Loaded lazily by Layout__Interop via ESM import with cache-busting.

export function getWindowWidth() {
    return window.innerWidth;
}

// Subscribes to window resize and invokes OnWindowWidthChanged on the .NET ref.
// Returns a handle whose dispose() removes the listener.
export function watchWindowWidth(dotNetRef) {
    function handler() {
        dotNetRef.invokeMethodAsync('OnWindowWidthChanged', window.innerWidth);
    }
    window.addEventListener('resize', handler);
    return {
        dispose: function () {
            window.removeEventListener('resize', handler);
        }
    };
}
