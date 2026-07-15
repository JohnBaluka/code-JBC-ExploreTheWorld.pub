using JBC.ExploreTheWorld.AL.BlazorLib;
using Radzen;

namespace JBC.ExploreTheWorld.RazorTests;

public class BlazorTestBase : BunitContext
{
    public BlazorTestBase()
    {
        Services.AddScoped<DialogService>();
        Services.AddScoped<NotificationService>();
        Services.AddScoped<TooltipService>();
        Services.AddScoped<ContextMenuService>();
        Services.AddSingleton<WatcherEvent_AppService>();
        Services.AddSingleton(new Layout_AppService());
        Services.AddSingleton(new DbProvider_AppService { ProviderName = "SqliteDb" });
        Services.AddScoped(_ => new Mock<OfficeExport_AppService__Interface>().Object);

        JSInterop.Mode = JSRuntimeMode.Loose;
    }
}
