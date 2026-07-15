using System.Diagnostics;

namespace JBC.ExploreTheWorld.WebAppTests;

public sealed class WebAppServerFixture : IAsyncLifetime
{
    private Process? _process;

    private static readonly string AppPath =
        Environment.GetEnvironmentVariable("ETW_WEBAPP_PATH")
        ?? FindAppExecutable("JBC.ExploreTheWorld.AL.BlazorWebApp.exe");

    public async Task InitializeAsync()
    {
        var baseUrl = Environment.GetEnvironmentVariable("ETW_BASE_URL") ?? "https://localhost:7000";

        if (await IsServerRunningAsync(baseUrl))
            return;

        var startInfo = new ProcessStartInfo
        {
            FileName               = AppPath,
            UseShellExecute        = false,
            RedirectStandardOutput = true,
            RedirectStandardError  = true,
            CreateNoWindow         = true
        };
        startInfo.Environment["ASPNETCORE_URLS"]        = "https://localhost:7000;http://localhost:7001";
        startInfo.Environment["ASPNETCORE_ENVIRONMENT"] = "Development";

        _process = new Process { StartInfo = startInfo };
        _process.OutputDataReceived += (_, _) => { };
        _process.ErrorDataReceived  += (_, _) => { };
        _process.Start();
        _process.BeginOutputReadLine();
        _process.BeginErrorReadLine();

        var deadline = DateTime.UtcNow.AddSeconds(30);
        while (DateTime.UtcNow < deadline)
        {
            if (await IsServerRunningAsync(baseUrl))
                return;
            await Task.Delay(500);
        }

        throw new TimeoutException($"Web app did not become ready within 30 seconds at {baseUrl}.");
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

    private static string FindAppExecutable(string exeName)
    {
        var repoRoot = Path.GetFullPath(
            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\..\..\..\"));

        var candidates = new[]
        {
            Path.Combine(repoRoot, "src", "AL.BlazorWebApp", "bin", "Debug",   "net10.0", exeName),
            Path.Combine(repoRoot, "src", "AL.BlazorWebApp", "bin", "Release", "net10.0", exeName)
        };

        return candidates.FirstOrDefault(File.Exists) ?? exeName;
    }
}
