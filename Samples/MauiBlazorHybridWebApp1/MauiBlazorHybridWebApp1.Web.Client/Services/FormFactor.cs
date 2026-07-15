using MauiBlazorHybridWebApp1.Shared.Services;

namespace MauiBlazorHybridWebApp1.Web.Client.Services
{
    public class FormFactor : IFormFactor
    {
        public string GetFormFactor()
        {
            return "WebAssembly";
        }

        public string GetPlatform()
        {
            return Environment.OSVersion.ToString();
        }
    }
}
