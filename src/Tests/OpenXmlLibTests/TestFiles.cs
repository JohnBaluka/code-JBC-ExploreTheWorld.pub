namespace JBC.ExploreTheWorld.OpenXmlLibTests;

internal static class TestFiles
{
    // Smallest valid PNG (1×1 transparent pixel) — used as a stand-in flag image.
    internal static byte[] TinyPng => Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==");

    internal static string NewTempPath(string extension) =>
        Path.Combine(Path.GetTempPath(), $"ETW_OpenXmlLibTests_{Guid.NewGuid():N}{extension}");

    internal static void Delete(params string[] paths)
    {
        foreach (var path in paths)
        {
            try { if (File.Exists(path)) File.Delete(path); } catch { }
        }
    }
}
