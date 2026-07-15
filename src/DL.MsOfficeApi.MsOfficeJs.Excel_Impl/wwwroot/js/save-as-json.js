// Collects the Office.js workbook object graph for Save-As-JSON. The returned dataJson
// matches the MsExcelJs/ExcelWorkbookJs_Row classes; the C# interop maps it to the
// canonical JBC.ExploreTheWorld.DL.MsOfficeApi.MsExcel entities so all writers (VBA, NetOffice,
// OpenXML, web add-in) emit the same schema.
export async function getWorkbookData() {
    const result = { dataJson: "", error: "", fileName: getDocumentFileName() };
    try {
        await Excel.run(async (context) => {
            const workbook = context.workbook;
            workbook.load("name");

            const sheets = workbook.worksheets;
            sheets.load("items/name,items/position,items/visibility");
            await context.sync();

            // Document properties require ExcelApi 1.7; ignore when unavailable. Loaded in
            // its own batch after the sync above so a failed sync here cannot discard the
            // workbook/sheets loads.
            let props = null;
            try {
                const properties = workbook.properties;
                properties.load("title,subject,author,keywords,comments,lastAuthor,revisionNumber,creationDate,category");
                await context.sync();
                props = {
                    title: properties.title || "",
                    subject: properties.subject || "",
                    author: properties.author || "",
                    keywords: properties.keywords || "",
                    comments: properties.comments || "",
                    lastAuthor: properties.lastAuthor || "",
                    revisionNumber: (properties.revisionNumber ?? "").toString(),
                    creationDate: properties.creationDate ? String(properties.creationDate) : "",
                    category: properties.category || ""
                };
            } catch (e) { /* document properties not available on this host */ }

            const sheetsData = [];
            for (const sheet of sheets.items) {
                const usedRange = sheet.getUsedRangeOrNullObject();
                usedRange.load("values,formulas,rowCount,columnCount,address");
                await context.sync();

                if (!usedRange.isNullObject) {
                    sheetsData.push({
                        name:        sheet.name,
                        position:    sheet.position,
                        visibility:  sheet.visibility || "Visible",
                        address:     usedRange.address,
                        rowCount:    usedRange.rowCount,
                        columnCount: usedRange.columnCount,
                        values:      usedRange.values.map(row => row.map(cellString)),
                        formulas:    usedRange.formulas.map(row => row.map(cellString))
                    });
                } else {
                    sheetsData.push({
                        name: sheet.name, position: sheet.position, visibility: sheet.visibility || "Visible",
                        address: "", rowCount: 0, columnCount: 0, values: [], formulas: []
                    });
                }
            }

            const data = { name: workbook.name, properties: props, sheets: sheetsData };
            result.dataJson = JSON.stringify(data);
        });
    } catch (error) {
        result.error = error.message || String(error);
    }
    return result;
}

function cellString(value) {
    return (value === undefined || value === null) ? "" : String(value);
}

// Returns the current document's file name (e.g. "ETW_CountriesNow.xlsx"), or "" when the
// document has not been saved / the URL is unavailable. Used as the Save-As-JSON default.
function getDocumentFileName() {
    try {
        const url = Office.context && Office.context.document && Office.context.document.url;
        if (url) {
            const name = url.split(/[\\/]/).pop();
            if (name) return decodeURIComponent(name);
        }
    } catch (e) { /* document url not available */ }
    return "";
}

export function downloadJson(json, fileName) {
    const blob = new Blob([json], { type: "application/json" });
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement("a");
    a.href     = url;
    a.download = fileName;
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 5000);
}
