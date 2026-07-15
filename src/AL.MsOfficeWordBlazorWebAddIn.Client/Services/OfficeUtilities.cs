using System.Runtime.InteropServices.JavaScript;
using System.Runtime.Versioning;

namespace ExploreTheWorld.AL.MsOfficeWordBlazorWebAddIn.Client.Services;

[SupportedOSPlatform("browser")]
public static partial class OfficeUtilities
{
    private static bool _isImported = false;

    public static async Task EnsureImportedAsync()
    {
        if (!_isImported)
        {
            try
            {
                await JSHost.ImportAsync("SharedUtils", "/scripts/SharedUtils.js");
                Console.WriteLine("Imported SharedUtils module");
                _isImported = true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error importing SharedUtils module: {ex.Message}");
                throw;
            }
        }
    }

    public static async Task<bool> IsRunningInHostAsync()
    {
        await EnsureImportedAsync();
        return await IsRunningInHostInternal();
    }

    [JSImport("IsRunningInHost", "SharedUtils")]
    private static partial Task<bool> IsRunningInHostInternal();
}
