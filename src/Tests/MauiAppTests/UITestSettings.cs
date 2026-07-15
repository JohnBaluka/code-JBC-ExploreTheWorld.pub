using System.Reflection;
using Xunit.Abstractions;

namespace JBC.ExploreTheWorld.MauiAppTests;

internal static class UITestSettings
{
    internal static bool ScreenshotsEnabled =>
        Environment.GetEnvironmentVariable("ETW_TEST_SCREENSHOTS") is not "false";

    // Set ETW_TEST_SETTLE_MS to override. The BlazorWebView (WebView2) paints
    // asynchronously after the window handle exists, so screenshots taken
    // immediately after launch are blank.
    internal static int SettleDelayMs =>
        int.TryParse(Environment.GetEnvironmentVariable("ETW_TEST_SETTLE_MS"), out var ms) ? ms : 4000;

    internal const int WindowWidth  = 1907;
    internal const int WindowHeight = 945;

    internal static readonly string OutputRoot = GetOutputRoot();

    private static string GetOutputRoot()
    {
        var repoRoot = Path.GetFullPath(
            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\..\..\..\"));
        return Path.Combine(repoRoot, "TestResults");
    }

    internal static string GetTestFolder(ITestOutputHelper output)
    {
        const string jbcPrefix = "JBC.ExploreTheWorld.";
        var assemblyName = Assembly.GetExecutingAssembly().GetName().Name ?? "";
        var projectName  = assemblyName.StartsWith(jbcPrefix) ? assemblyName[jbcPrefix.Length..] : assemblyName;

        var displayName = GetDisplayName(output);
        var subPrefix   = jbcPrefix + projectName + ".";
        var subPath     = displayName.StartsWith(subPrefix) ? displayName[subPrefix.Length..] : displayName;

        var segments = subPath.Split('.');
        return Path.Combine(new[] { OutputRoot, projectName }.Concat(segments).ToArray());
    }

    private static string GetDisplayName(ITestOutputHelper output)
    {
        try
        {
            var field = output.GetType()
                .GetFields(BindingFlags.NonPublic | BindingFlags.Instance)
                .FirstOrDefault(f => typeof(ITest).IsAssignableFrom(f.FieldType));
            if (field?.GetValue(output) is ITest test)
                return test.DisplayName;
        }
        catch { }
        return "unknown";
    }
}
