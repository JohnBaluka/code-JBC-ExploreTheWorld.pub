// Appends the loaded CountriesNow rows to the active Word document as a table
// (Flag, Country, ISO2, ISO3) with an inline flag image per row, via the Office.js API.
// payloadJson: { "countries": [ { "country", "iso2", "iso3", "flagUrl" }, ... ] }

const FLAG_HEIGHT_PT = 16;

export async function insertCountries(payloadJson) {
    const result = { inserted: false, count: 0, flagCount: 0, error: "" };
    try {
        const data = JSON.parse(payloadJson || "{}");
        const countries = data.countries || [];
        if (countries.length === 0) {
            result.error = "No countries to export. Load the countries first.";
            return result;
        }

        // Flag PNGs are fetched in the browser (parallel) with their pixel dimensions.
        const flags = await fetchFlags(countries);

        await Word.run(async (context) => {
            const body = context.document.body;

            const values = [["Flag", "Country", "ISO2", "ISO3"]];
            for (const c of countries) {
                values.push(["", c.country || "", c.iso2 || "", c.iso3 || ""]);
            }

            // Append a blank paragraph + the table at the very end of the document.
            const endRange = body.getRange(Word.RangeLocation.end);
            const table = endRange.insertTable(values.length, 4, Word.InsertLocation.after, values);
            table.rows.getFirst().font.bold = true;
            await context.sync();

            // Insert the flag image into the first column of each data row.
            let flagCount = 0;
            for (let i = 0; i < countries.length; i++) {
                const flag = flags[i];
                if (!flag || !flag.base64) continue;
                const cell = table.getCell(i + 1, 0);
                const picture = cell.body.insertInlinePictureFromBase64(flag.base64, Word.InsertLocation.start);
                // Set the width explicitly from the flag's real pixel aspect ratio so it is not
                // distorted (setting height alone does not rescale width reliably).
                picture.height = FLAG_HEIGHT_PT;
                picture.width = flag.w > 0 && flag.h > 0
                    ? FLAG_HEIGHT_PT * (flag.w / flag.h)
                    : FLAG_HEIGHT_PT * 1.5;
                flagCount++;
            }
            await context.sync();

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
