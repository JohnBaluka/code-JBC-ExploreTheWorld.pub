export async function getPresentationInfo() {
    const result = {
        errorMessage: "",
        title: "",
        slideCount: 0,
        slideWidth: 0,
        slideHeight: 0,
        author: ""
    };

    try {
        // Presentation.properties requires PowerPointApi 1.7 and Presentation.pageSetup
        // (slide dimensions) requires PowerPointApi 1.10; the manifest minimum is 1.5.
        const supportsProperties = Office.context.requirements.isSetSupported("PowerPointApi", "1.7");
        const supportsPageSetup  = Office.context.requirements.isSetSupported("PowerPointApi", "1.10");

        await PowerPoint.run(async (context) => {
            const presentation = context.presentation;
            presentation.load("title");

            const slides = presentation.slides;
            slides.load("items/id");

            let properties = null;
            if (supportsProperties) {
                properties = presentation.properties;
                properties.load("author");
            }

            let pageSetup = null;
            if (supportsPageSetup) {
                pageSetup = presentation.pageSetup;
                pageSetup.load("slideWidth,slideHeight");
            }

            await context.sync();

            result.title = presentation.title || "";
            result.slideCount = slides.items.length;
            result.author = properties ? (properties.author || "") : "";
            result.slideWidth = pageSetup ? (pageSetup.slideWidth ?? 0) : 0;
            result.slideHeight = pageSetup ? (pageSetup.slideHeight ?? 0) : 0;
        });
    } catch (error) {
        result.errorMessage = error.message || String(error);
    }

    return result;
}
