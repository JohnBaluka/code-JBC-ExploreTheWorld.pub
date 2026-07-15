// Collects the Office.js document object graph for Save-As-JSON. The returned dataJson
// matches the MsWordJs/WordDocumentJs_Row classes; the C# interop maps it to the
// canonical JBC.ExploreTheWorld.DL.MsOfficeApi.MsWord entities so all writers (VBA, NetOffice,
// OpenXML, web add-in) emit the same schema.
export async function getDocumentData() {
    const result = { dataJson: "", error: "", fileName: getDocumentFileName() };
    try {
        await Word.run(async (context) => {
            const doc = context.document;

            // Document properties.
            let props = null;
            try {
                const properties = doc.properties;
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

            // Paragraphs with formatting.
            const paragraphs = doc.body.paragraphs;
            paragraphs.load("items/text,items/style,items/styleBuiltIn,items/alignment,items/leftIndent,items/rightIndent,items/firstLineIndent,items/spaceBefore,items/spaceAfter,items/lineSpacing");
            await context.sync();

            const paragraphsData = paragraphs.items.map(p => ({
                text: p.text || "",
                style: p.style || p.styleBuiltIn || "",
                alignment: p.alignment || null,
                leftIndent: numberOrNull(p.leftIndent),
                rightIndent: numberOrNull(p.rightIndent),
                firstLineIndent: numberOrNull(p.firstLineIndent),
                spaceBefore: numberOrNull(p.spaceBefore),
                spaceAfter: numberOrNull(p.spaceAfter),
                lineSpacing: numberOrNull(p.lineSpacing)
            }));

            // Tables.
            let tablesData = [];
            try {
                const tables = doc.body.tables;
                tables.load("items/rowCount,items/values");
                await context.sync();
                tablesData = tables.items.map(t => ({
                    rowCount: t.rowCount || 0,
                    values: (t.values || []).map(row => row.map(v => (v === undefined || v === null) ? "" : String(v)))
                }));
            } catch (e) { /* tables not available on this host */ }

            // Content controls.
            let contentControlsData = [];
            try {
                const contentControls = doc.contentControls;
                contentControls.load("items/title,items/tag,items/text");
                await context.sync();
                contentControlsData = contentControls.items.map(cc => ({
                    title: cc.title || "",
                    tag: cc.tag || "",
                    text: cc.text || ""
                }));
            } catch (e) { /* content controls not available */ }

            // Inline pictures with base64 image bytes.
            let inlinePicturesData = [];
            try {
                const pictures = doc.body.inlinePictures;
                pictures.load("items/width,items/height,items/altTextDescription");
                await context.sync();

                const base64Results = pictures.items.map(p => p.getBase64ImageSrc());
                await context.sync();

                inlinePicturesData = pictures.items.map((p, i) => ({
                    base64: base64Results[i] ? base64Results[i].value : null,
                    width: numberOrNull(p.width),
                    height: numberOrNull(p.height),
                    altTextDescription: p.altTextDescription || ""
                }));
            } catch (e) { /* inline pictures not available */ }

            const data = {
                properties: props,
                paragraphs: paragraphsData,
                tables: tablesData,
                contentControls: contentControlsData,
                inlinePictures: inlinePicturesData
            };
            result.dataJson = JSON.stringify(data);
        });
    } catch (error) {
        result.error = error.message || String(error);
    }
    return result;
}

function numberOrNull(value) {
    return (value === undefined || value === null || isNaN(value)) ? null : value;
}

// Returns the current document's file name (e.g. "ETW_CountriesNow.docx"), or "" when the
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
