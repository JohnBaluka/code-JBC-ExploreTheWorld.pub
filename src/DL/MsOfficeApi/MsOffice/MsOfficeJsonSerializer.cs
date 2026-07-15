using System;
using System.IO;
using System.Text;
using System.Text.Json;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsOffice
{
    // Canonical JSON serializer for the MsOffice/MsPowerPoint/MsExcel/MsWord entities.
    // The settings are chosen so every writer (VBA, NetOffice, OpenXML, Office.js interop)
    // produces byte-identical output for the same object graph:
    //   - 2-space indent, CRLF line endings (matches the VBA TextStream writers)
    //   - default encoder (matches the customized json_Encode in VBA JsonConverter.bas)
    //   - null property values are written explicitly
    public static class MsOfficeJsonSerializer
    {
        private static readonly JsonSerializerOptions _options = CreateOptions();

        public static JsonSerializerOptions CreateOptions()
        {
            return new JsonSerializerOptions
            {
                WriteIndented = true,
                NewLine = "\r\n",
                AllowTrailingCommas = true,
                ReadCommentHandling = JsonCommentHandling.Skip,
            };
        }

        public static string Serialize<T>(T value)
        {
            return JsonSerializer.Serialize(value, _options);
        }

        public static T? Deserialize<T>(string json)
        {
            return JsonSerializer.Deserialize<T>(json, _options);
        }

        // Writes UTF-8 without BOM and a trailing newline (matching the VBA writers).
        public static void WriteToFile<T>(T value, string outputFilePath)
        {
            string json = Serialize(value) + "\r\n";
            File.WriteAllText(outputFilePath, json, new UTF8Encoding(false));
        }

        public static T? ReadFromFile<T>(string inputFilePath)
        {
            return Deserialize<T>(File.ReadAllText(inputFilePath));
        }
    }
}
