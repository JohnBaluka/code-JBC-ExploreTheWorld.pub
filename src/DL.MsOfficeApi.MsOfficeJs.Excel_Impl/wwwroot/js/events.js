let _dotNetRef = null;
let _eventContexts = [];

function _timestamp() {
    return new Date().toISOString().substring(11, 23);
}

function _notify(eventName) {
    if (_dotNetRef) {
        _dotNetRef.invokeMethodAsync("OnEventLogged", eventName, _timestamp());
    }
}

function _onWorksheetActivated(args)   { _notify(`WorksheetActivated: ${args.worksheetId}`); }
function _onWorksheetDeactivated(args) { _notify(`WorksheetDeactivated: ${args.worksheetId}`); }
function _onWorksheetAdded(args)       { _notify(`WorksheetAdded: ${args.worksheetId}`); }
function _onWorksheetDeleted(args)     { _notify(`WorksheetDeleted: ${args.worksheetId}`); }
function _onSelectionChanged(args)     { _notify(`SelectionChanged: ${args.address}`); }

// Worksheet collection events (ExcelApi 1.7); the manifest minimum is 1.7.
const _collectionHandlerMap = {
    worksheetActivated:   { prop: "onActivated",   fn: _onWorksheetActivated },
    worksheetDeactivated: { prop: "onDeactivated", fn: _onWorksheetDeactivated },
    worksheetAdded:       { prop: "onAdded",       fn: _onWorksheetAdded },
    worksheetDeleted:     { prop: "onDeleted",     fn: _onWorksheetDeleted },
};

export async function startWatching(dotNetRef, eventKeys) {
    _dotNetRef = dotNetRef;
    _eventContexts = [];

    try {
        await Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;

            for (const key of eventKeys) {
                const entry = _collectionHandlerMap[key];
                if (!entry) continue;
                const eventHandlers = sheets[entry.prop];
                if (!eventHandlers) {
                    console.warn(`[Events] Excel.WorksheetCollection.${entry.prop} unavailable in this environment`);
                    continue;
                }
                _eventContexts.push(eventHandlers.add(entry.fn));
            }

            if (eventKeys.includes("selectionChanged")) {
                if (sheets.onSelectionChanged) {
                    // Collection-level event (ExcelApi 1.9) fires on every worksheet.
                    _eventContexts.push(sheets.onSelectionChanged.add(_onSelectionChanged));
                } else {
                    // Fallback (ExcelApi 1.7): only the worksheet active right now is watched.
                    const activeSheet = sheets.getActiveWorksheet();
                    if (activeSheet.onSelectionChanged) {
                        _eventContexts.push(activeSheet.onSelectionChanged.add(_onSelectionChanged));
                    } else {
                        console.warn("[Events] Excel.Worksheet.onSelectionChanged unavailable in this environment");
                    }
                }
            }

            await context.sync();
        });
    } catch (error) {
        throw new Error("startWatching failed: " + (error.message || String(error)));
    }
}

export async function stopWatching() {
    try {
        if (_eventContexts.length > 0) {
            await Excel.run(async (context) => {
                for (const ctx of _eventContexts) {
                    ctx.remove();
                }
                await context.sync();
            });
        }
    } catch (error) {
        console.warn("[Events] stopWatching cleanup error:", error);
    } finally {
        _eventContexts = [];
        _dotNetRef = null;
    }
}
