export async function getWorkbookInfo() {
    const result = {
        errorMessage: "",
        workbookName: "",
        activeSheetName: "",
        sheetCount: 0,
        usedRangeAddress: ""
    };

    try {
        await Excel.run(async (context) => {
            const workbook = context.workbook;
            workbook.load("name");

            const activeSheet = workbook.worksheets.getActiveWorksheet();
            activeSheet.load("name");

            const sheets = workbook.worksheets;
            sheets.load("items/name");

            await context.sync();

            result.workbookName = workbook.name || "";
            result.activeSheetName = activeSheet.name || "";
            result.sheetCount = sheets.items.length;

            const usedRange = activeSheet.getUsedRangeOrNullObject();
            usedRange.load("address");
            await context.sync();

            result.usedRangeAddress = usedRange.isNullObject ? "(empty)" : (usedRange.address || "");
        });
    } catch (error) {
        result.errorMessage = error.message || String(error);
    }

    return result;
}
