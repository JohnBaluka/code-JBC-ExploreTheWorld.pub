using System.Diagnostics;

namespace JBC.ExploreTheWorld.OqtaneTests;

/// <summary>
/// Starts the local Oqtane host (oqtane.framework/Oqtane.Server) if it is not already
/// running. Requires the installed local site: SQL Server LocalDB with the Oqtane-ETW
/// database, the ETW pages created (SQL/002) and the ExploreTheWorld theme set as the
/// site default (SQL/004).
/// </summary>
public sealed class OqtaneServerFixture : IAsyncLifetime
{
    private Process? _process;

    private static readonly string ServerDirectory =
        Environment.GetEnvironmentVariable("ETW_OQTANE_PATH")
        ?? FindServerDirectory();

    public async Task InitializeAsync()
    {
        var baseUrl = Environment.GetEnvironmentVariable("ETW_OQTANE_URL") ?? "http://localhost:44357";

        if (await IsServerRunningAsync(baseUrl))
            return;

        var exePath = Path.Combine(ServerDirectory, "bin", "Debug", "net10.0", "Oqtane.Server.exe");
        if (!File.Exists(exePath))
            throw new FileNotFoundException(
                $"Oqtane.Server.exe not found at {exePath}. Build oqtane.framework/Oqtane.Server first.");

        var startInfo = new ProcessStartInfo
        {
            FileName               = exePath,
            // Oqtane resolves appsettings.json and the Data folder (LocalDB mdf) from the
            // content root, which defaults to the working directory — use the project dir.
            WorkingDirectory       = ServerDirectory,
            UseShellExecute        = false,
            RedirectStandardOutput = true,
            RedirectStandardError  = true,
            CreateNoWindow         = true
        };
        startInfo.Environment["ASPNETCORE_URLS"]        = baseUrl;
        startInfo.Environment["ASPNETCORE_ENVIRONMENT"] = "Development";

        _process = new Process { StartInfo = startInfo };
        _process.OutputDataReceived += (_, _) => { };
        _process.ErrorDataReceived  += (_, _) => { };
        _process.Start();
        _process.BeginOutputReadLine();
        _process.BeginErrorReadLine();

        var deadline = DateTime.UtcNow.AddSeconds(60);
        while (DateTime.UtcNow < deadline)
        {
            if (await IsServerRunningAsync(baseUrl))
                return;
            await Task.Delay(500);
        }

        throw new TimeoutException($"Oqtane did not become ready within 60 seconds at {baseUrl}.");
    }

    public Task DisposeAsync()
    {
        if (_process is { HasExited: false })
        {
            _process.Kill(entireProcessTree: true);
            _process.Dispose();
        }
        return Task.CompletedTask;
    }

    private static async Task<bool> IsServerRunningAsync(string baseUrl)
    {
        try
        {
            using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(2) };
            var response = await http.GetAsync(baseUrl);
            return (int)response.StatusCode < 500;
        }
        catch
        {
            return false;
        }
    }

    private static string FindServerDirectory()
    {
        var repoRoot = Path.GetFullPath(
            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\..\..\..\"));
        return Path.Combine(repoRoot, "oqtane.framework", "Oqtane.Server");
    }
}
