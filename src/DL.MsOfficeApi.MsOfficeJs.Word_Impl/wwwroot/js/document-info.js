export async function getDocumentInfo() {
    const result = {
        errorMessage: "",
        title: "",
        author: "",
        wordCount: 0,
        paragraphCount: 0,
        pageCount: 0,
        revision: 0
    };

    try {
        await Word.run(async (context) => {
            // Word.DocumentProperties has no wordCount/paragraphCount/pageCount; counts are
            // derived from the body below. The JS API does not expose a page count at all,
            // so pageCount stays 0.
            const properties = context.document.properties;
            properties.load("title,author,revisionNumber");

            const body = context.document.body;
            body.load("text");

            const paragraphs = body.paragraphs;
            paragraphs.load("items");

            await context.sync();

            result.title = properties.title || "";
            result.author = properties.author || "";
            result.revision = parseInt(properties.revisionNumber, 10) || 0;
            result.wordCount = ((body.text || "").match(/\S+/g) || []).length;
            result.paragraphCount = paragraphs.items.length;
        });
    } catch (error) {
        result.errorMessage = error.message || String(error);
    }

    return result;
}
