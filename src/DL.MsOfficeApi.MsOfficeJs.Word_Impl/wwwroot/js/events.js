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

function _onDocumentSelectionChanged() { _notify("SelectionChanged"); }
function _onParagraphAdded()           { _notify("ParagraphAdded"); }
function _onParagraphChanged()         { _notify("ParagraphChanged"); }
function _onParagraphDeleted()         { _notify("ParagraphDeleted"); }
function _onContentControlAdded()      { _notify("ContentControlAdded"); }
function _onContentControlDeleted()    { _notify("ContentControlDeleted"); }
function _onContentControlEntered()    { _notify("ContentControlEntered"); }
function _onContentControlExited()     { _notify("ContentControlExited"); }

// Events that exist on Word.Document. Deleted/entered/exited are NOT document-level
// events — they live on each Word.ContentControl instance (WordApi 1.5) and are
// registered separately below.
const _documentHandlerMap = {
    paragraphAdded:      { prop: "onParagraphAdded",      fn: _onParagraphAdded },
    paragraphChanged:    { prop: "onParagraphChanged",    fn: _onParagraphChanged },
    paragraphDeleted:    { prop: "onParagraphDeleted",    fn: _onParagraphDeleted },
    contentControlAdded: { prop: "onContentControlAdded", fn: _onContentControlAdded },
};

const _contentControlHandlerMap = {
    contentControlDeleted: { prop: "onDeleted", fn: _onContentControlDeleted },
    contentControlEntered: { prop: "onEntered", fn: _onContentControlEntered },
    contentControlExited:  { prop: "onExited",  fn: _onContentControlExited },
};

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
        await Word.run(async (context) => {
            for (const key of eventKeys) {
                const entry = _documentHandlerMap[key];
                if (!entry) continue;
                const eventHandlers = context.document[entry.prop];
                if (!eventHandlers) {
                    console.warn(`[Events] Word.Document.${entry.prop} unavailable in this environment`);
                    continue;
                }
                _eventContexts.push(eventHandlers.add(entry.fn));
            }

            // Register the per-instance content control events on the controls that exist
            // now; controls added after watching starts are not watched.
            const contentControlKeys = eventKeys.filter(key => _contentControlHandlerMap[key]);
            if (contentControlKeys.length > 0) {
                const contentControls = context.document.contentControls;
                contentControls.load("items");
                await context.sync();

                for (const contentControl of contentControls.items) {
                    for (const key of contentControlKeys) {
                        const entry = _contentControlHandlerMap[key];
                        const eventHandlers = contentControl[entry.prop];
                        if (!eventHandlers) {
                            console.warn(`[Events] Word.ContentControl.${entry.prop} unavailable in this environment`);
                            continue;
                        }
                        _eventContexts.push(eventHandlers.add(entry.fn));
                    }
                }
            }

            await context.sync();
        });

        if (eventKeys.includes("selectionChanged")) {
            // Word.Document has no production onSelectionChanged event; the Office common
            // API covers selection changes on all hosts.
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
            await Word.run(async (context) => {
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
