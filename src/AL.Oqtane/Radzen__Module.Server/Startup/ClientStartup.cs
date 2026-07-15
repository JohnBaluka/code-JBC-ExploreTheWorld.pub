using Microsoft.Extensions.DependencyInjection;
using Oqtane.Services;

namespace JBC.ExploreTheWorld.AL.Oqtane.Radzen
{
    public class ClientStartup : IClientStartup
    {
        public void ConfigureServices(IServiceCollection services)
        {
            services.AddScoped<global::Radzen.DialogService>();
            services.AddScoped<global::Radzen.NotificationService>();
            services.AddScoped<global::Radzen.TooltipService>();
            services.AddScoped<global::Radzen.ContextMenuService>();
        }
    }
}
