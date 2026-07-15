using JBC.ExploreTheWorld.CL;

namespace JBC.ExploreTheWorld.AL.SaveAsJson.ConsoleApp
{
    /// <summary>
    /// Command line options mirroring the Watcher forms' Save-As-JSON controls: the Office
    /// host (inferred from the source file extension unless overridden), the JSON write
    /// method (Direct/Interop/Dynamic/NetOffice/OpenXml), and the output path (defaulting to the
    /// source file name with ".json" appended, like the watcher forms).
    /// </summary>
    public sealed class SaveAsJsonOptions
    {
        public string Type { get; private set; } = string.Empty;
        public string Method { get; private set; } = "NetOffice";
        public string SourcePath { get; private set; } = string.Empty;
        public string OutputPath { get; private set; } = string.Empty;
        public bool ShowHelp { get; private set; }
        public string? Error { get; private set; }

        private static readonly string[] Types = ["Word", "Excel", "PowerPoint"];
        private static readonly string[] Methods = ["Direct", "Interop", "Dynamic", "NetOffice", "OpenXml"];

        public const string Usage = """
            ExploreTheWorld SaveAsJson — writes an Office document's canonical JSON representation.

            Usage:
              JBC.ExploreTheWorld.AL.SaveAsJson.ConsoleApp --source <document> [options]

            Options (the same choices as the Watcher forms):
              --source <path>                                Office document to read (required).
              --method <Direct|Interop|Dynamic|NetOffice|OpenXml>
                                                             JSON write method. Default: NetOffice.
                                                             Direct runs the ExploreTheWorld VBA macro
                                                             inside the Office host and requires the
                                                             macro template to be loaded. Interop uses
                                                             the Microsoft.Office.Interop PIAs.
              --type <Word|Excel|PowerPoint>                 Office host. Default: inferred from the
                                                             source file extension.
              --output <path>                                Output .json path. Default: the source
                                                             file name with the method inserted and
                                                             ".json" appended (e.g. Report-Direct.pptx.json).
              --help                                         Show this help.
            """;

        public static SaveAsJsonOptions Parse(string[] args)
        {
            var options = new SaveAsJsonOptions();

            for (int i = 0; i < args.Length; i++)
            {
                switch (args[i].ToLowerInvariant())
                {
                    case "--help":
                    case "-h":
                    case "/?":
                        options.ShowHelp = true;
                        return options;
                    case "--type":
                        options.Type = NextValue(args, ref i, options) ?? options.Type;
                        break;
                    case "--method":
                        options.Method = NextValue(args, ref i, options) ?? options.Method;
                        break;
                    case "--source":
                        options.SourcePath = NextValue(args, ref i, options) ?? options.SourcePath;
                        break;
                    case "--output":
                        options.OutputPath = NextValue(args, ref i, options) ?? options.OutputPath;
                        break;
                    default:
                        options.Error = $"Unknown option: {args[i]}";
                        return options;
                }

                if (options.Error != null) return options;
            }

            if (string.IsNullOrWhiteSpace(options.SourcePath))
            {
                options.Error = "--source is required.";
                return options;
            }

            options.SourcePath = Path.GetFullPath(options.SourcePath);
            if (!File.Exists(options.SourcePath))
            {
                options.Error = $"Source file not found: {options.SourcePath}";
                return options;
            }

            if (string.IsNullOrWhiteSpace(options.Type))
            {
                options.Type = Path.GetExtension(options.SourcePath).ToLowerInvariant() switch
                {
                    ".docx" or ".docm" or ".dotx" or ".dotm" or ".doc" => "Word",
                    ".xlsx" or ".xlsm" or ".xltx" or ".xltm" or ".xls" => "Excel",
                    ".pptx" or ".pptm" or ".potx" or ".potm" or ".ppt" => "PowerPoint",
                    _ => string.Empty,
                };

                if (options.Type.Length == 0)
                {
                    options.Error = "The Office host could not be inferred from the source extension — pass --type.";
                    return options;
                }
            }

            options.Type = MatchChoice(Types, options.Type, "--type", options);
            options.Method = MatchChoice(Methods, options.Method, "--method", options);
            if (options.Error != null) return options;

            if (string.IsNullOrWhiteSpace(options.OutputPath))
            {
                options.OutputPath = SaveAsJson_Helper.BuildDefaultPath(options.SourcePath, $"ETW_Ms{options.Type}", options.Method);
            }

            options.OutputPath = Path.GetFullPath(options.OutputPath);
            return options;
        }

        private static string? NextValue(string[] args, ref int i, SaveAsJsonOptions options)
        {
            if (i + 1 >= args.Length)
            {
                options.Error = $"Missing value for {args[i]}.";
                return null;
            }

            i++;
            return args[i];
        }

        // Choices are matched case-insensitively and normalized to their canonical casing.
        private static string MatchChoice(string[] choices, string value, string optionName, SaveAsJsonOptions options)
        {
            var match = choices.FirstOrDefault(c => string.Equals(c, value, StringComparison.OrdinalIgnoreCase));
            if (match == null)
            {
                options.Error = $"Invalid value '{value}' for {optionName}. Expected: {string.Join(" | ", choices)}.";
                return value;
            }

            return match;
        }
    }
}
