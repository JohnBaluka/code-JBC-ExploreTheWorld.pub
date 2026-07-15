using System.Diagnostics;

namespace JBC.ExploreTheWorld.OfficeWebAddinTests;

public sealed class OfficeWebAddinServerFixture : IAsyncLifetime
{
    private readonly List<Process> _processes = new();

    private static readonly string WordAppPath =
        Environment.GetEnvironmentVariable("ETW_WORD_ADDIN_PATH")
        ?? FindAppExecutable("AL.MsOfficeWordBlazorWebAddIn", "ExploreTheWorld.AL.MsOfficeWordBlazorWebAddIn.exe");

    private static readonly string ExcelAppPath =
        Environment.GetEnvironmentVariable("ETW_EXCEL_ADDIN_PATH")
        ?? FindAppExecutable("AL.MsOfficeExcelBlazorWebAddIn", "ExploreTheWorld.AL.MsOfficeExcelBlazorWebAddIn.exe");

    private static readonly string PowerPointAppPath =
        Environment.GetEnvironmentVariable("ETW_PPT_ADDIN_PATH")
        ?? FindAppExecutable("AL.MsOfficePowerPointBlazorWebAddIn", "ExploreTheWorld.AL.MsOfficePowerPointBlazorWebAddIn.exe");

    public async Task InitializeAsync()
    {
        await StartServerAsync(WordAppPath,       "https://localhost:7100", "https://localhost:7100;http://localhost:5100");
        await StartServerAsync(ExcelAppPath,      "https://localhost:7101", "https://localhost:7101;http://localhost:5101");
        await StartServerAsync(PowerPointAppPath, "https://localhost:7102", "https://localhost:7102;http://localhost:5102");
    }

    public Task DisposeAsync()
    {
        foreach (var process in _processes)
        {
            if (process is { HasExited: false })
            {
                process.Kill(entireProcessTree: true);
                process.Dispose();
            }
        }
        return Task.CompletedTask;
    }

    private async Task StartServerAsync(string exePath, string baseUrl, string aspNetUrls)
    {
        if (await IsServerRunningAsync(baseUrl))
            return;

        var startInfo = new ProcessStartInfo
        {
            FileName               = exePath,
            UseShellExecute        = false,
            RedirectStandardOutput = true,
            RedirectStandardError  = true,
            CreateNoWindow         = true
        };
        startInfo.Environment["ASPNETCORE_URLS"]        = aspNetUrls;
        startInfo.Environment["ASPNETCORE_ENVIRONMENT"] = "Development";

        var process = new Process { StartInfo = startInfo };
        process.OutputDataReceived += (_, _) => { };
        process.ErrorDataReceived  += (_, _) => { };
        process.Start();
        process.BeginOutputReadLine();
        process.BeginErrorReadLine();
        _processes.Add(process);

        var deadline = DateTime.UtcNow.AddSeconds(30);
        while (DateTime.UtcNow < deadline)
        {
            if (await IsServerRunningAsync(baseUrl))
                return;
            await Task.Delay(500);
        }

        throw new TimeoutException($"Add-in server did not become ready within 30 seconds at {baseUrl}.");
    }

    private static async Task<bool> IsServerRunningAsync(string baseUrl)
    {
        try
        {
            using var handler = new HttpClientHandler
            {
                ServerCertificateCustomValidationCallback = (_, _, _, _) => true
            };
            using var http = new HttpClient(handler) { Timeout = TimeSpan.FromSeconds(2) };
            var response = await http.GetAsync(baseUrl);
            return (int)response.StatusCode < 500;
        }
        catch
        {
            return false;
        }
    }

    private static string FindAppExecutable(string projectFolder, string exeName)
    {
        var repoRoot = Path.GetFullPath(
            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\..\..\..\"));

        var candidates = new[]
        {
            Path.Combine(repoRoot, "src", projectFolder, "bin", "Debug",   "net10.0", exeName),
            Path.Combine(repoRoot, "src", projectFolder, "bin", "Release", "net10.0", exeName)
        };

        return candidates.FirstOrDefault(File.Exists) ?? exeName;
    }
}
