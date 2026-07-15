using System.Collections.Generic;

namespace JBC.ExploreTheWorld.AL.WinFormsLib
{
    /// <summary>
    /// Hard-coded repository of MS Office application events.
    /// Mirrors the MsOfficeEvents table in ExploreTheWorld.accdb — not stored in a database.
    /// </summary>
    public static class MsOfficeEvents_Repo
    {
        public static IReadOnlyList<MsOfficeEvent_Record> All { get; } = new List<MsOfficeEvent_Record>
        {
            // ── Word — Application ───────────────────────────────────────────────────
            new MsOfficeEvent_Record { Name = "Quit",                              Category = "Application",    Word = true,  Excel = false, PowerPoint = false },

            // ── Word — Document ──────────────────────────────────────────────────────
            new MsOfficeEvent_Record { Name = "NewDocument",                       Category = "Document",       Word = true,  Excel = false, PowerPoint = false },
            new MsOfficeEvent_Record { Name = "DocumentOpen",                      Category = "Document",       Word = true,  Excel = false, PowerPoint = false },
            new MsOfficeEvent_Record { Name = "DocumentBeforeClose",               Category = "Document",       Word = true,  Excel = false, PowerPoint = false },
            new MsOfficeEvent_Record { Name = "DocumentBeforePrint",               Category = "Document",       Word = true,  Excel = false, PowerPoint = false },
            new MsOfficeEvent_Record { Name = "DocumentBeforeSave",                Category = "Document",       Word = true,  Excel = false, PowerPoint = false },
            new MsOfficeEvent_Record { Name = "DocumentChange",                    Category = "Document",       Word = true,  Excel = false, PowerPoint = false },
            new MsOfficeEvent_Record { Name = "DocumentSync",                      Category = "Document",       Word = true,  Excel = false, PowerPoint = false },

            // ── Word — E-Postage ─────────────────────────────────────────────────────
            new MsOfficeEvent_Record { Name = "EPostageInsert",                    Category = "E-Postage",      Word = true,  Excel = false, PowerPoint = false },
            new MsOfficeEvent_Record { Name = "EPostageInsertEx",                  Category = "E-Postage",      Word = true,  Excel = false, PowerPoint = false },
            new MsOfficeEvent_Record { Name = "EPostagePropertyDialog",            Category = "E-Postage",      Word = true,  Excel = false, PowerPoint = false },

            // ── Word — Mail Merge ────────────────────────────────────────────────────
            new MsOfficeEvent_Record { Name = "MailMergeAfterMerge",               Category = "Mail Merge",     Word = true,  Excel = false, PowerPoint = false },
            new MsOfficeEvent_Record { Name = "MailMergeAfterRecordMerge",         Category = "Mail Merge",     Word = true,  Excel = false, PowerPoint = false },
            new MsOfficeEvent_Record { Name = "MailMergeBeforeMerge",              Category = "Mail Merge",     Word = true,  Excel = false, PowerPoint = false },
            new MsOfficeEvent_Record { Name = "MailMergeBeforeRecordMerge",        Category = "Mail Merge",     Word = true,  Excel = false, PowerPoint = false },
            new MsOfficeEvent_Record { Name = "MailMergeDataSourceLoad",           Category = "Mail Merge",     Word = true,  Excel = false, PowerPoint = false },
            new MsOfficeEvent_Record { Name = "MailMergeDataSourceValidate",       Category = "Mail Merge",     Word = true,  Excel = false, PowerPoint = false },
            new MsOfficeEvent_Record { Name = "MailMergeDataSourceValidate2",      Category = "Mail Merge",     Word = true,  Excel = false, PowerPoint = false },
            new MsOfficeEvent_Record { Name = "MailMergeWizardSendToCustom",       Category = "Mail Merge",     Word = true,  Excel = false, PowerPoint = false },
            new MsOfficeEvent_Record { Name = "MailMergeWizardStateChange",        Category = "Mail Merge",     Word = true,  Excel = false, PowerPoint = false },

            // ── Word — XML ───────────────────────────────────────────────────────────
            new MsOfficeEvent_Record { Name = "XMLSelectionChange",                Category = "XML",            Word = true,  Excel = false, PowerPoint = false },
            new MsOfficeEvent_Record { Name = "XMLValidationError",                Category = "XML",            Word = true,  Excel = false, PowerPoint = false },

            // ── Excel — Calculation ──────────────────────────────────────────────────
            new MsOfficeEvent_Record { Name = "AfterCalculate",                    Category = "Calculation",    Word = false, Excel = true,  PowerPoint = false },

            // ── Excel — Chart ────────────────────────────────────────────────────────
            new MsOfficeEvent_Record { Name = "WorkbookNewChart",                  Category = "Chart",          Word = false, Excel = true,  PowerPoint = false },

            // ── Excel — Data ─────────────────────────────────────────────────────────
            new MsOfficeEvent_Record { Name = "WorkbookRowsetComplete",            Category = "Data",           Word = false, Excel = true,  PowerPoint = false },

            // ── Excel — Data Model ───────────────────────────────────────────────────
            new MsOfficeEvent_Record { Name = "WorkbookModelChange",               Category = "Data Model",     Word = false, Excel = true,  PowerPoint = false },

            // ── Excel — PivotTable ───────────────────────────────────────────────────
            new MsOfficeEvent_Record { Name = "SheetPivotTableAfterValueChange",   Category = "PivotTable",     Word = false, Excel = true,  PowerPoint = false },
            new MsOfficeEvent_Record { Name = "SheetPivotTableBeforeAllocateChanges", Category = "PivotTable",  Word = false, Excel = true,  PowerPoint = false },
            new MsOfficeEvent_Record { Name = "SheetPivotTableBeforeCommitChanges",Category = "PivotTable",     Word = false, Excel = true,  PowerPoint = false },
            new MsOfficeEvent_Record { Name = "SheetPivotTableBeforeDiscardChanges", Category = "PivotTable",   Word = false, Excel = true,  PowerPoint = false },
            new MsOfficeEvent_Record { Name = "SheetPivotTableUpdate",             Category = "PivotTable",     Word = false, Excel = true,  PowerPoint = false },
            new MsOfficeEvent_Record { Name = "WorkbookPivotTableCloseConnection", Category = "PivotTable",     Word = false, Excel = true,  PowerPoint = false },
            new MsOfficeEvent_Record { Name = "WorkbookPivotTableOpenConnection",  Category = "PivotTable",     Word = false, Excel = true,  PowerPoint = false },

            // ── Excel — Table ────────────────────────────────────────────────────────
            new MsOfficeEvent_Record { Name = "SheetTableUpdate",                  Category = "Table",          Word = false, Excel = true,  PowerPoint = false },

            // ── Excel — Workbook ─────────────────────────────────────────────────────
            new MsOfficeEvent_Record { Name = "NewWorkbook",                       Category = "Workbook",       Word = false, Excel = true,  PowerPoint = false },
            new MsOfficeEvent_Record { Name = "WorkbookActivate",                  Category = "Workbook",       Word = false, Excel = true,  PowerPoint = false },
            new MsOfficeEvent_Record { Name = "WorkbookAfterSave",                 Category = "Workbook",       Word = false, Excel = true,  PowerPoint = false },
            new MsOfficeEvent_Record { Name = "WorkbookBeforeClose",               Category = "Workbook",       Word = false, Excel = true,  PowerPoint = false },
            new MsOfficeEvent_Record { Name = "WorkbookBeforePrint",               Category = "Workbook",       Word = false, Excel = true,  PowerPoint = false },
            new MsOfficeEvent_Record { Name = "WorkbookBeforeSave",                Category = "Workbook",       Word = false, Excel = true,  PowerPoint = false },
            new MsOfficeEvent_Record { Name = "WorkbookDeactivate",                Category = "Workbook",       Word = false, Excel = true,  PowerPoint = false },
            new MsOfficeEvent_Record { Name = "WorkbookOpen",                      Category = "Workbook",       Word = false, Excel = true,  PowerPoint = false },
            new MsOfficeEvent_Record { Name = "WorkbookSync",                      Category = "Workbook",       Word = false, Excel = true,  PowerPoint = false },

            // ── Excel — Workbook/Add-in ──────────────────────────────────────────────
            new MsOfficeEvent_Record { Name = "WorkbookAddinInstall",              Category = "Workbook/Add-in",Word = false, Excel = true,  PowerPoint = false },
            new MsOfficeEvent_Record { Name = "WorkbookAddinUninstall",            Category = "Workbook/Add-in",Word = false, Excel = true,  PowerPoint = false },

            // ── Excel — Worksheet ────────────────────────────────────────────────────
            new MsOfficeEvent_Record { Name = "SheetActivate",                     Category = "Worksheet",      Word = false, Excel = true,  PowerPoint = false },
            new MsOfficeEvent_Record { Name = "SheetBeforeDelete",                 Category = "Worksheet",      Word = false, Excel = true,  PowerPoint = false },
            new MsOfficeEvent_Record { Name = "SheetBeforeDoubleClick",            Category = "Worksheet",      Word = false, Excel = true,  PowerPoint = false },
            new MsOfficeEvent_Record { Name = "SheetBeforeRightClick",             Category = "Worksheet",      Word = false, Excel = true,  PowerPoint = false },
            new MsOfficeEvent_Record { Name = "SheetCalculate",                    Category = "Worksheet",      Word = false, Excel = true,  PowerPoint = false },
            new MsOfficeEvent_Record { Name = "SheetChange",                       Category = "Worksheet",      Word = false, Excel = true,  PowerPoint = false },
            new MsOfficeEvent_Record { Name = "SheetDeactivate",                   Category = "Worksheet",      Word = false, Excel = true,  PowerPoint = false },
            new MsOfficeEvent_Record { Name = "SheetFollowHyperlink",              Category = "Worksheet",      Word = false, Excel = true,  PowerPoint = false },
            new MsOfficeEvent_Record { Name = "SheetLensGalleryRenderComplete",    Category = "Worksheet",      Word = false, Excel = true,  PowerPoint = false },
            new MsOfficeEvent_Record { Name = "SheetSelectionChange",              Category = "Worksheet",      Word = false, Excel = true,  PowerPoint = false },
            new MsOfficeEvent_Record { Name = "WorkbookNewSheet",                  Category = "Worksheet",      Word = false, Excel = true,  PowerPoint = false },

            // ── Excel — XML ──────────────────────────────────────────────────────────
            new MsOfficeEvent_Record { Name = "WorkbookAfterXmlExport",            Category = "XML",            Word = false, Excel = true,  PowerPoint = false },
            new MsOfficeEvent_Record { Name = "WorkbookAfterXmlImport",            Category = "XML",            Word = false, Excel = true,  PowerPoint = false },
            new MsOfficeEvent_Record { Name = "WorkbookBeforeXmlExport",           Category = "XML",            Word = false, Excel = true,  PowerPoint = false },
            new MsOfficeEvent_Record { Name = "WorkbookBeforeXmlImport",           Category = "XML",            Word = false, Excel = true,  PowerPoint = false },

            // ── PowerPoint — Presentation ────────────────────────────────────────────
            new MsOfficeEvent_Record { Name = "AfterNewPresentation",              Category = "Presentation",   Word = false, Excel = false, PowerPoint = true  },
            new MsOfficeEvent_Record { Name = "AfterPresentationOpen",             Category = "Presentation",   Word = false, Excel = false, PowerPoint = true  },
            new MsOfficeEvent_Record { Name = "ColorSchemeChanged",                Category = "Presentation",   Word = false, Excel = false, PowerPoint = true  },
            new MsOfficeEvent_Record { Name = "NewPresentation",                   Category = "Presentation",   Word = false, Excel = false, PowerPoint = true  },
            new MsOfficeEvent_Record { Name = "PresentationBeforeClose",           Category = "Presentation",   Word = false, Excel = false, PowerPoint = true  },
            new MsOfficeEvent_Record { Name = "PresentationBeforeSave",            Category = "Presentation",   Word = false, Excel = false, PowerPoint = true  },
            new MsOfficeEvent_Record { Name = "PresentationClose",                 Category = "Presentation",   Word = false, Excel = false, PowerPoint = true  },
            new MsOfficeEvent_Record { Name = "PresentationCloseFinal",            Category = "Presentation",   Word = false, Excel = false, PowerPoint = true  },
            new MsOfficeEvent_Record { Name = "PresentationOpen",                  Category = "Presentation",   Word = false, Excel = false, PowerPoint = true  },
            new MsOfficeEvent_Record { Name = "PresentationPrint",                 Category = "Presentation",   Word = false, Excel = false, PowerPoint = true  },
            new MsOfficeEvent_Record { Name = "PresentationSave",                  Category = "Presentation",   Word = false, Excel = false, PowerPoint = true  },
            new MsOfficeEvent_Record { Name = "PresentationSync",                  Category = "Presentation",   Word = false, Excel = false, PowerPoint = true  },

            // ── PowerPoint — Shape ───────────────────────────────────────────────────
            new MsOfficeEvent_Record { Name = "AfterShapeSizeChange",              Category = "Shape",          Word = false, Excel = false, PowerPoint = true  },

            // ── PowerPoint — Slide ───────────────────────────────────────────────────
            new MsOfficeEvent_Record { Name = "AfterDragDropOnSlide",              Category = "Slide",          Word = false, Excel = false, PowerPoint = true  },
            new MsOfficeEvent_Record { Name = "PresentationNewSlide",              Category = "Slide",          Word = false, Excel = false, PowerPoint = true  },
            new MsOfficeEvent_Record { Name = "SlideSelectionChanged",             Category = "Slide",          Word = false, Excel = false, PowerPoint = true  },

            // ── PowerPoint — Slide Show ──────────────────────────────────────────────
            new MsOfficeEvent_Record { Name = "SlideShowBegin",                    Category = "Slide Show",     Word = false, Excel = false, PowerPoint = true  },
            new MsOfficeEvent_Record { Name = "SlideShowEnd",                      Category = "Slide Show",     Word = false, Excel = false, PowerPoint = true  },
            new MsOfficeEvent_Record { Name = "SlideShowNextBuild",                Category = "Slide Show",     Word = false, Excel = false, PowerPoint = true  },
            new MsOfficeEvent_Record { Name = "SlideShowNextClick",                Category = "Slide Show",     Word = false, Excel = false, PowerPoint = true  },
            new MsOfficeEvent_Record { Name = "SlideShowNextSlide",                Category = "Slide Show",     Word = false, Excel = false, PowerPoint = true  },
            new MsOfficeEvent_Record { Name = "SlideShowOnNext",                   Category = "Slide Show",     Word = false, Excel = false, PowerPoint = true  },
            new MsOfficeEvent_Record { Name = "SlideShowOnPrevious",               Category = "Slide Show",     Word = false, Excel = false, PowerPoint = true  },

            // ── Shared — Protected View (Word + Excel + PowerPoint) ──────────────────
            new MsOfficeEvent_Record { Name = "ProtectedViewWindowActivate",       Category = "Protected View", Word = true,  Excel = true,  PowerPoint = true  },
            new MsOfficeEvent_Record { Name = "ProtectedViewWindowBeforeClose",    Category = "Protected View", Word = true,  Excel = true,  PowerPoint = true  },
            new MsOfficeEvent_Record { Name = "ProtectedViewWindowBeforeEdit",     Category = "Protected View", Word = true,  Excel = true,  PowerPoint = true  },
            new MsOfficeEvent_Record { Name = "ProtectedViewWindowDeactivate",     Category = "Protected View", Word = true,  Excel = true,  PowerPoint = true  },
            new MsOfficeEvent_Record { Name = "ProtectedViewWindowOpen",           Category = "Protected View", Word = true,  Excel = true,  PowerPoint = true  },

            // ── Excel-only Protected View ────────────────────────────────────────────
            new MsOfficeEvent_Record { Name = "ProtectedViewWindowResize",         Category = "Protected View", Word = false, Excel = true,  PowerPoint = false },

            // ── Word-only Protected View ─────────────────────────────────────────────
            new MsOfficeEvent_Record { Name = "ProtectedViewWindowSize",           Category = "Protected View", Word = true,  Excel = false, PowerPoint = false },

            // ── Shared — Window (Word + Excel + PowerPoint) ──────────────────────────
            new MsOfficeEvent_Record { Name = "WindowActivate",                    Category = "Window",         Word = true,  Excel = true,  PowerPoint = true  },
            new MsOfficeEvent_Record { Name = "WindowDeactivate",                  Category = "Window",         Word = true,  Excel = true,  PowerPoint = true  },

            // ── Shared — Window (Word + PowerPoint) ──────────────────────────────────
            new MsOfficeEvent_Record { Name = "WindowBeforeDoubleClick",           Category = "Window",         Word = true,  Excel = false, PowerPoint = true  },
            new MsOfficeEvent_Record { Name = "WindowBeforeRightClick",            Category = "Window",         Word = true,  Excel = false, PowerPoint = true  },
            new MsOfficeEvent_Record { Name = "WindowSelectionChange",             Category = "Window",         Word = true,  Excel = false, PowerPoint = true  },

            // ── Excel-only Window ────────────────────────────────────────────────────
            new MsOfficeEvent_Record { Name = "WindowResize",                      Category = "Window",         Word = false, Excel = true,  PowerPoint = false },

            // ── Word-only Window ─────────────────────────────────────────────────────
            new MsOfficeEvent_Record { Name = "WindowSize",                        Category = "Window",         Word = true,  Excel = false, PowerPoint = false },
        };

        /// <summary>Returns a cloned, mutable list of Word events for use within a single form instance.</summary>
        public static List<MsOfficeEvent_Record> ForWord()
        {
            var result = new List<MsOfficeEvent_Record>();
            foreach (var e in All)
            {
                if (e.Word)
                    result.Add(new MsOfficeEvent_Record
                    {
                        Name       = e.Name,
                        Category   = e.Category,
                        Word       = e.Word,
                        Excel      = e.Excel,
                        PowerPoint = e.PowerPoint,
                        Log        = e.Log
                    });
            }
            return result;
        }

        /// <summary>Returns a cloned, mutable list of Excel events for use within a single form instance.</summary>
        public static List<MsOfficeEvent_Record> ForExcel()
        {
            var result = new List<MsOfficeEvent_Record>();
            foreach (var e in All)
            {
                if (e.Excel)
                    result.Add(new MsOfficeEvent_Record
                    {
                        Name       = e.Name,
                        Category   = e.Category,
                        Word       = e.Word,
                        Excel      = e.Excel,
                        PowerPoint = e.PowerPoint,
                        Log        = e.Log
                    });
            }
            return result;
        }

        /// <summary>Returns a cloned, mutable list of PowerPoint events for use within a single form instance.</summary>
        public static List<MsOfficeEvent_Record> ForPowerPoint()
        {
            var result = new List<MsOfficeEvent_Record>();
            foreach (var e in All)
            {
                if (e.PowerPoint)
                    result.Add(new MsOfficeEvent_Record
                    {
                        Name       = e.Name,
                        Category   = e.Category,
                        Word       = e.Word,
                        Excel      = e.Excel,
                        PowerPoint = e.PowerPoint,
                        Log        = e.Log
                    });
            }
            return result;
        }
    }
}
