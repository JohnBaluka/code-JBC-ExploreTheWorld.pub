using System;
using System.IO;
using NetOffice.WordApi;
using DynamicWriters = JBC.ExploreTheWorld.DL.MsOfficeApi.Dynamic_Impl.JsonWriters;
using InteropWriters = JBC.ExploreTheWorld.DL.MsOfficeApi.Interop_Impl.JsonWriters;
using NetOfficeWriters = JBC.ExploreTheWorld.DL.MsOfficeApi.NetOffice_Impl.JsonWriters;
using OpenXmlWriters = JBC.ExploreTheWorld.DL.MsOfficeApi.OpenXml_Impl.JsonWriters;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi_Impl
{
    // Thin wrapper over the DL "Save as JSON" writers so the watcher forms keep their
    // existing call sites. Both paths emit the canonical
    // JBC.ExploreTheWorld.DL.MsOfficeApi.MsWord entity schema (same as the VBA writers).
    internal static class MsWord_JsonWriter
    {
        /// <summary>Writes a full JSON representation of the Word document using the NetOffice COM object model.</summary>
        internal static void WriteNetOffice(Document doc, string filePath, Action<string> log)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (string.IsNullOrEmpty(filePath)) throw new ArgumentException("filePath is required.", nameof(filePath));

            NetOfficeWriters.MsWordJsonWriter.WriteDocumentToJsonFile(doc, filePath, null, log);
        }

        /// <summary>Writes a full JSON representation of the Word document using the late-bound dynamic COM writers (DL.MsOfficeApi.Dynamic_Impl).</summary>
        internal static void WriteDynamic(Document doc, string filePath, Action<string> log)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (string.IsNullOrEmpty(filePath)) throw new ArgumentException("filePath is required.", nameof(filePath));

            DynamicWriters.MsWordJsonWriter.WriteDocumentComToJsonFile(doc.UnderlyingObject, filePath, null, log);
        }

        /// <summary>Writes a full JSON representation of the Word document using the strongly-typed Microsoft.Office.Interop COM object model (DL.MsOfficeApi.Interop_Impl).</summary>
        internal static void WriteInterop(Document doc, string filePath, Action<string> log)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (string.IsNullOrEmpty(filePath)) throw new ArgumentException("filePath is required.", nameof(filePath));

            // The watcher holds a NetOffice wrapper; unwrap to the raw COM object and cast to the
            // Interop interface (a QueryInterface on the same live Word.Document COM object).
            var interopDoc = (Microsoft.Office.Interop.Word.Document)doc.UnderlyingObject;
            InteropWriters.MsWordJsonWriter.WriteDocumentToJsonFile(interopDoc, filePath, null, log);
        }

        /// <summary>Writes a JSON representation of the .docx file at <paramref name="sourceDocxPath"/> using OpenXML (no COM required).</summary>
        internal static void WriteOpenXml(string sourceDocxPath, string outputJsonFilePath, Action<string> log)
        {
            if (string.IsNullOrEmpty(sourceDocxPath)) throw new ArgumentException("sourceDocxPath is required.", nameof(sourceDocxPath));
            if (string.IsNullOrEmpty(outputJsonFilePath)) throw new ArgumentException("outputJsonFilePath is required.", nameof(outputJsonFilePath));

            OpenXmlWriters.MsWordJsonWriter.WriteDocumentFileToJson(sourceDocxPath, outputJsonFilePath, null, log);
        }

        /// <summary>
        /// Writes the running app's active document to JSON using OpenXML. OpenXML reads the file
        /// on disk while Word holds a lock on it, so the document is saved, closed, read, and
        /// reopened. Cloud (AutoSave/OneDrive) documents have an https path OpenXML cannot read,
        /// and Word has no SaveCopyAs — those report a clear error instead.
        /// </summary>
        internal static void WriteOpenXmlFromRunningApp(NetOffice.WordApi.Application wordApp, string outputJsonFilePath, Action<string> log)
        {
            if (wordApp == null) throw new ArgumentNullException(nameof(wordApp));

            var doc = wordApp.ActiveDocument
                ?? throw new InvalidOperationException("No active document.");

            string documentFolder = string.Empty;
            try { documentFolder = doc.Path; } catch { }
            if (string.IsNullOrEmpty(documentFolder))
                throw new InvalidOperationException(
                    "The active document has never been saved — save it first, or use the NetOffice method.");

            string fullName = doc.FullName;

            if (fullName.StartsWith("http", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    "OpenXML cannot read cloud (AutoSave) documents — use the NetOffice or Direct method, or save the document to a local folder.");
            }

            log("Closing active document for OpenXML read...");
            try { doc.Save(); } catch { }
            doc.Close(false);
            try
            {
                WriteOpenXml(fullName, outputJsonFilePath, log);
            }
            finally
            {
                log("Reopening document...");
                wordApp.Documents.Open(fullName);
            }
        }
    }
}
