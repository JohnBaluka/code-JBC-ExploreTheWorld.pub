namespace JBC.ExploreTheWorld.AL.ExportData.ConsoleApp
{
    /// <summary>
    /// Command line options mirroring the Countries API form's export controls:
    /// export type (Word/Excel/PowerPoint), export method (Interop/Dynamic/NetOffice/OpenXml),
    /// database provider, and output file path.
    /// </summary>
    public sealed class ExportDataOptions
    {
        public string Type { get; private set; } = "Word";
        public string Method { get; private set; } = "OpenXml";
        public string Provider { get; private set; } = "InMemoryDb";
        public string FilePath { get; private set; } = string.Empty;
        public bool ShowHelp { get; private set; }
        public string? Error { get; private set; }

        private static readonly string[] Types = ["Word", "Excel", "PowerPoint"];
        private static readonly string[] Methods = ["Interop", "Dynamic", "NetOffice", "OpenXml"];
        private static readonly string[] Providers = ["InMemoryDb", "AccessDb", "SqliteDb", "SqlServerDb"];

        public const string Usage = """
            ExploreTheWorld ExportData — exports the countries list to an Office document.

            Usage:
              JBC.ExploreTheWorld.AL.ExportData.ConsoleApp [options]

            Options (the same choices as the Countries API form):
              --type <Word|Excel|PowerPoint>                     Export type. Default: Word.
              --method <Interop|Dynamic|NetOffice|OpenXml>       Export method. Default: OpenXml.
                                                                 Interop/NetOffice/Dynamic require Microsoft Office.
              --provider <InMemoryDb|AccessDb|SqliteDb|SqlServerDb>
                                                                 Database provider (DB-backed cache with
                                                                 API fallback). Default: InMemoryDb.
              --file <path>                                      Output file path. Default:
                                                                 ETW_CountriesNow[-Access]-{method}.{docx|xlsx|pptx}
                                                                 ("Access" is inserted when --provider is AccessDb).
              --help                                             Show this help.
            """;

        public static ExportDataOptions Parse(string[] args)
        {
            var options = new ExportDataOptions();

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
                    case "--provider":
                        options.Provider = NextValue(args, ref i, options) ?? options.Provider;
                        break;
                    case "--file":
                        options.FilePath = NextValue(args, ref i, options) ?? options.FilePath;
                        break;
                    default:
                        options.Error = $"Unknown option: {args[i]}";
                        return options;
                }

                if (options.Error != null) return options;
            }

            options.Type = MatchChoice(Types, options.Type, "--type", options);
            options.Method = MatchChoice(Methods, options.Method, "--method", options);
            options.Provider = MatchChoice(Providers, options.Provider, "--provider", options);

            if (options.Error != null) return options;

            if (string.IsNullOrWhiteSpace(options.FilePath))
            {
                var extension = options.Type switch
                {
                    "Excel" => "xlsx",
                    "PowerPoint" => "pptx",
                    _ => "docx",
                };
                var fromAccessDb = string.Equals(options.Provider, "AccessDb", StringComparison.OrdinalIgnoreCase);
                options.FilePath = JBC.ExploreTheWorld.CL.MsOfficeExportName_Helper.BuildFileName(
                    "ETW_CountriesNow", options.Method, fromAccessDb, extension);
            }

            options.FilePath = Path.GetFullPath(options.FilePath);
            return options;
        }

        private static string? NextValue(string[] args, ref int i, ExportDataOptions options)
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
        private static string MatchChoice(string[] choices, string value, string optionName, ExportDataOptions options)
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
