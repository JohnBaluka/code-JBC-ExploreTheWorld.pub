using System.Reflection;
using Xunit.Abstractions;

namespace JBC.ExploreTheWorld.OfficeWebAddinTests;

internal static class UITestSettings
{
    // Set ETW_TEST_SCREENSHOTS=false to disable
    internal static bool ScreenshotsEnabled =>
        Environment.GetEnvironmentVariable("ETW_TEST_SCREENSHOTS") is not "false";

    // Set ETW_TEST_VIDEO=false to disable
    internal static bool VideoEnabled =>
        Environment.GetEnvironmentVariable("ETW_TEST_VIDEO") is not "false";

    // Set ETW_TEST_HEADLESS=true to run headless (default: visible)
    internal static bool Headless =>
        Environment.GetEnvironmentVariable("ETW_TEST_HEADLESS") is "true";

    // Set ETW_TEST_SETTLE_MS to override. Blazor keeps rendering after the network
    // goes idle, so screenshots wait this long for the page to finish painting.
    internal static int SettleDelayMs =>
        int.TryParse(Environment.GetEnvironmentVariable("ETW_TEST_SETTLE_MS"), out var ms) ? ms : 1500;

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

        var displayName     = GetDisplayName(output);
        var mainProjectName = projectName.Replace("._netF", "");
        var subPrefix       = jbcPrefix + mainProjectName + ".";
        var subPath         = displayName.StartsWith(subPrefix) ? displayName[subPrefix.Length..] : displayName;

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
