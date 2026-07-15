// Builds a "Country Slides"-style deck in the active presentation via the Office.js API:
// a title slide ("Explore the World" + count) followed by one centered slide per country
// (large name, centered flag, ISO2/ISO3). The flag is drawn as a rectangle whose fill is the
// base64 PNG (ShapeCollection has no addImageFromBase64) — the rectangle is sized to the flag's
// real pixel aspect ratio so it is not distorted. When the presentation is the default single
// starter slide, that slide is reused as the title slide so there is no leading blank.
// payloadJson: { "countries": [ { "country", "iso2", "iso3", "flagUrl" }, ... ] }

// PowerPoint on the web default slide size (Widescreen 16:9 = 13.333in x 7.5in @ 72pt/in).
const SLIDE_W = 960;
const SLIDE_H = 540;

export async function insertCountries(payloadJson) {
    const result = { inserted: false, count: 0, flagCount: 0, flagError: "", error: "" };
    try {
        const data = JSON.parse(payloadJson || "{}");
        const countries = data.countries || [];
        if (countries.length === 0) {
            result.error = "No countries to export. Load the countries first.";
            return result;
        }

        const flags = await fetchFlags(countries);

        await PowerPoint.run(async (context) => {
            const slides = context.presentation.slides;
            slides.load("items/id");
            await context.sync();

            // A fresh deck opens with a single starter slide — reuse it as the title slide so the
            // export has no leading blank. If the deck already has content, append a title slide.
            const startCount = slides.items.length;
            const reuseFirst = startCount === 1;
            const toAdd = reuseFirst ? countries.length : countries.length + 1;

            for (let i = 0; i < toAdd; i++) slides.add();
            await context.sync();
            slides.load("items/id");
            await context.sync();

            // Title slide.
            const titleSlide = reuseFirst ? slides.items[0] : slides.items[startCount];
            addCenteredTextBox(titleSlide, "Explore the World",
                SLIDE_W * 0.06, SLIDE_H * 0.32, SLIDE_W * 0.88, SLIDE_H * 0.22, 40, true);
            addCenteredTextBox(titleSlide, `${countries.length} countries`,
                SLIDE_W * 0.06, SLIDE_H * 0.56, SLIDE_W * 0.88, SLIDE_H * 0.12, 20, false);

            // Country slides — name + ISO codes.
            const firstCountry = reuseFirst ? startCount : startCount + 1;
            for (let i = 0; i < countries.length; i++) {
                const c = countries[i];
                const slide = slides.items[firstCountry + i];
                if (!slide) continue;
                addCenteredTextBox(slide, c.country || "",
                    SLIDE_W * 0.06, SLIDE_H * 0.06, SLIDE_W * 0.88, SLIDE_H * 0.20, 36, true);
                const codes = [c.iso2, c.iso3].filter(Boolean).join("      ");
                addCenteredTextBox(slide, codes,
                    SLIDE_W * 0.06, SLIDE_H * 0.80, SLIDE_W * 0.88, SLIDE_H * 0.12, 20, false);
            }
            await context.sync();

            result.count = countries.length;
            result.inserted = true;

            // Country slides — flags (best-effort; a failure here leaves the slides intact).
            try {
                let flagCount = 0;
                for (let i = 0; i < countries.length; i++) {
                    const flag = flags[i];
                    if (!flag || !flag.base64) continue;
                    const slide = slides.items[firstCountry + i];
                    if (!slide) continue;

                    const h = SLIDE_H * 0.42;
                    // Size to the flag's real aspect ratio so the fill is not stretched; center it.
                    const w = flag.w > 0 && flag.h > 0 ? h * (flag.w / flag.h) : h * 1.5;
                    const shape = slide.shapes.addGeometricShape(
                        PowerPoint.GeometricShapeType.rectangle,
                        { left: (SLIDE_W - w) / 2, top: SLIDE_H * 0.30, width: w, height: h });
                    shape.fill.setImage(flag.base64);
                    try { shape.lineFormat.visible = false; } catch { /* border stays */ }
                    flagCount++;
                }
                await context.sync();
                result.flagCount = flagCount;
            } catch (flagEx) {
                result.flagError = flagEx.message || String(flagEx);
            }
        });
    } catch (error) {
        result.error = error.message || String(error);
    }
    return result;
}

// Adds a horizontally-centered text box. Font size/weight and alignment are best-effort
// (wrapped so older Office.js builds without paragraphFormat still insert the text).
function addCenteredTextBox(slide, text, left, top, width, height, fontSize, bold) {
    const shape = slide.shapes.addTextBox(text, { left, top, width, height });
    try {
        const range = shape.textFrame.textRange;
        range.font.size = fontSize;
        range.font.bold = bold;
        range.paragraphFormat.horizontalAlignment = "Center";
    } catch { /* text formatting not supported on this host — leave default */ }
    return shape;
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
