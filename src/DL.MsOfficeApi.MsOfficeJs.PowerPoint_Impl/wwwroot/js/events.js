let _dotNetRef = null;
let _eventContexts = [];
let _commonSelectionHandlerAdded = false;

function _timestamp() {
    return new Date().toISOString().substring(11, 23);
}

function _notify(eventName) {
    if (_dotNetRef) {
        _dotNetRef.invokeMethodAsync("OnEventLogged", eventName, _timestamp());
    }
}

function _onSlideSelectionChanged(args) { _notify(`SlideSelectionChanged: startIndex=${args.startSlideIndex}`); }
function _onDocumentSelectionChanged()  { _notify("SelectionChanged"); }

function _addSelectionHandlerAsync(handler) {
    return new Promise((resolve, reject) => {
        Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, handler, (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) resolve();
            else reject(new Error(asyncResult.error ? asyncResult.error.message : "addHandlerAsync failed"));
        });
    });
}

function _removeSelectionHandlerAsync(handler) {
    return new Promise((resolve, reject) => {
        Office.context.document.removeHandlerAsync(Office.EventType.DocumentSelectionChanged, { handler: handler }, (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) resolve();
            else reject(new Error(asyncResult.error ? asyncResult.error.message : "removeHandlerAsync failed"));
        });
    });
}

export async function startWatching(dotNetRef, eventKeys) {
    _dotNetRef = dotNetRef;
    _eventContexts = [];
    _commonSelectionHandlerAdded = false;

    try {
        if (eventKeys.includes("slideSelectionChanged")) {
            await PowerPoint.run(async (context) => {
                // Presentation.onSlideSelectionChanged is a preview API (PowerPointApi BETA)
                // and is undefined on production hosts.
                const eventHandlers = context.presentation.onSlideSelectionChanged;
                if (eventHandlers) {
                    _eventContexts.push(eventHandlers.add(_onSlideSelectionChanged));
                } else {
                    console.warn("[Events] PowerPoint.Presentation.onSlideSelectionChanged unavailable in this environment");
                }
                await context.sync();
            });
        }

        if (eventKeys.includes("selectionChanged")) {
            // The PowerPoint rich API has no production selection event; the Office common
            // API covers slide, shape, and text selection changes on all hosts.
            await _addSelectionHandlerAsync(_onDocumentSelectionChanged);
            _commonSelectionHandlerAdded = true;
        }
    } catch (error) {
        throw new Error("startWatching failed: " + (error.message || String(error)));
    }
}

export async function stopWatching() {
    try {
        if (_eventContexts.length > 0) {
            await PowerPoint.run(async (context) => {
                for (const ctx of _eventContexts) {
                    ctx.remove();
                }
                await context.sync();
            });
        }
        if (_commonSelectionHandlerAdded) {
            await _removeSelectionHandlerAsync(_onDocumentSelectionChanged);
        }
    } catch (error) {
        console.warn("[Events] stopWatching cleanup error:", error);
    } finally {
        _eventContexts = [];
        _commonSelectionHandlerAdded = false;
        _dotNetRef = null;
    }
}
