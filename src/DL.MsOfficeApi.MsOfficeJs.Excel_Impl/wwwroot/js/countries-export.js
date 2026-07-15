// Writes the loaded CountriesNow rows to the active worksheet starting at cell A1
// (columns: Flag, Country, ISO2, ISO3), then overlays a flag image per row when the
// host supports Excel image shapes (ExcelApi 1.9). Existing cells at A1 are overwritten.
// payloadJson: { "countries": [ { "country", "iso2", "iso3", "flagUrl" }, ... ] }

const ROW_HEIGHT_PT = 15;  // uniform height so the floating flag images line up with rows
const FLAG_COL_WIDTH_PT = 40;

export async function insertCountries(payloadJson) {
    const result = { inserted: false, count: 0, flagCount: 0, error: "" };
    try {
        const data = JSON.parse(payloadJson || "{}");
        const countries = data.countries || [];
        if (countries.length === 0) {
            result.error = "No countries to export. Load the countries first.";
            return result;
        }

        const flags = await fetchFlags(countries);
        const canImage = Office.context.requirements.isSetSupported("ExcelApi", "1.9");

        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const rowCount = countries.length + 1;

            const values = [["Flag", "Country", "ISO2", "ISO3"]];
            for (const c of countries) {
                values.push(["", c.country || "", c.iso2 || "", c.iso3 || ""]);
            }

            const dataRange = sheet.getRangeByIndexes(0, 0, rowCount, 4);
            dataRange.values = values;
            dataRange.format.rowHeight = ROW_HEIGHT_PT;

            const headerRange = sheet.getRangeByIndexes(0, 0, 1, 4);
            headerRange.format.font.bold = true;

            // Autofit the text columns; keep column A fixed for the flag images.
            sheet.getRangeByIndexes(0, 1, rowCount, 3).format.autofitColumns();
            sheet.getRangeByIndexes(0, 0, 1, 1).format.columnWidth = FLAG_COL_WIDTH_PT;

            await context.sync();

            let flagCount = 0;
            if (canImage) {
                const flagHeight = ROW_HEIGHT_PT - 3;
                for (let i = 0; i < countries.length; i++) {
                    const flag = flags[i];
                    if (!flag || !flag.base64) continue;
                    try {
                        const shape = sheet.shapes.addImage(flag.base64);
                        // Excel's lockAspectRatio does not rescale width when only height is set,
                        // so set the width explicitly from the flag's real pixel aspect ratio.
                        shape.lockAspectRatio = false;
                        shape.left = 2;
                        shape.top = (i + 1) * ROW_HEIGHT_PT + 1.5;
                        shape.height = flagHeight;
                        shape.width = flag.w > 0 && flag.h > 0
                            ? flagHeight * (flag.w / flag.h)
                            : flagHeight * 1.5;
                        flagCount++;
                    } catch {
                        // image insertion not available for this row — leave the cell blank
                    }
                }
                await context.sync();
            }

            result.count = countries.length;
            result.flagCount = flagCount;
            result.inserted = true;
        });
    } catch (error) {
        result.error = error.message || String(error);
    }
    return result;
}

// Fetches each country's flag PNG and returns { base64, w, h } (pixel size from the PNG header)
// per country, or null when there is no URL or the download fails.
async function fetchFlags(countries) {
    return await Promise.all(countries.map(async (c) => {
        if (!c.flagUrl) return null;
        try {
            const response = await fetch(c.flagUrl);
            if (!response.ok) return null;
            const buffer = await response.arrayBuffer();
            const size = readPngSize(buffer);
            return { base64: arrayBufferToBase64(buffer), w: size.w, h: size.h };
        } catch {
            return null;
        }
    }));
}

// Reads a PNG's pixel dimensions from its IHDR chunk (width @ byte 16, height @ 20, big-endian).
function readPngSize(buffer) {
    try {
        const dv = new DataView(buffer);
        const w = dv.getUint32(16);
        const h = dv.getUint32(20);
        if (w > 0 && h > 0) return { w, h };
    } catch { /* not a PNG / too short */ }
    return { w: 0, h: 0 };
}

function arrayBufferToBase64(buffer) {
    let binary = "";
    const bytes = new Uint8Array(buffer);
    const chunkSize = 0x8000;
    for (let i = 0; i < bytes.length; i += chunkSize) {
        binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunkSize));
    }
    return btoa(binary);
}
