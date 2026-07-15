// Collects the Office.js presentation object graph for Save-As-JSON. The returned
// dataJson matches the MsPowerPointJs/PowerPointPresentationJs_Row classes; the C#
// interop maps it to the canonical JBC.ExploreTheWorld.DL.MsOfficeApi.MsPowerPoint entities so all
// writers (VBA, NetOffice, OpenXML, web add-in) emit the same schema.
export async function getPresentationData() {
    const result = { dataJson: "", error: "", fileName: getDocumentFileName() };
    try {
        // Shape alt text requires PowerPointApi 1.10; the manifest minimum is 1.5.
        const supportsAltText = Office.context.requirements.isSetSupported("PowerPointApi", "1.10");

        await PowerPoint.run(async (context) => {
            const presentation = context.presentation;
            presentation.load("title");

            const slides = presentation.slides;
            slides.load("items/id,items/layout/name");
            await context.sync();

            const slidesData = [];
            for (let i = 0; i < slides.items.length; i++) {
                const slide = slides.items[i];

                const shapes = slide.shapes;
                let shapeProperties = "items/id,items/name,items/type,items/left,items/top,items/width,items/height";
                if (supportsAltText) {
                    shapeProperties += ",items/altTextDescription,items/altTextTitle";
                }
                shapes.load(shapeProperties);
                await context.sync();

                const shapesData = [];
                for (let j = 0; j < shapes.items.length; j++) {
                    const shape = shapes.items[j];

                    // Shape.textFrame throws for shapes without a text frame, which fails the
                    // whole batch — so text is probed with one sync per shape.
                    let text = null;
                    try {
                        const textRange = shape.textFrame.textRange;
                        textRange.load("text");
                        await context.sync();
                        text = textRange.text;
                    } catch (e) { /* shape has no text frame */ }

                    shapesData.push({
                        id: shape.id || "",
                        name: shape.name || "",
                        type: shape.type || "",
                        left: numberOrNull(shape.left),
                        top: numberOrNull(shape.top),
                        width: numberOrNull(shape.width),
                        height: numberOrNull(shape.height),
                        altTextDescription: supportsAltText ? (shape.altTextDescription || "") : null,
                        altTextTitle: supportsAltText ? (shape.altTextTitle || "") : null,
                        zOrderIndex: j + 1,
                        text: text
                    });
                }

                slidesData.push({
                    id: slide.id,
                    index: i + 1,
                    layoutName: slide.layout.name || null,
                    shapes: shapesData
                });
            }

            const data = { title: presentation.title || "", slides: slidesData };
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

// Returns the current document's file name (e.g. "ETW_CountriesNow.pptx"), or "" when the
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
