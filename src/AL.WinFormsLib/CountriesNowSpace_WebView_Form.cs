using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using JBC.ExploreTheWorld.AL.BlazorLib;
using Microsoft.AspNetCore.Components.WebView.WindowsForms;
using Microsoft.Extensions.DependencyInjection;

namespace JBC.ExploreTheWorld.AL.WinFormsLib
{
    public partial class CountriesNowSpace_WebView_Form : Form
    {
        public CountriesNowSpace_WebView_Form(ServiceProvider serviceProvider)
        {
            InitializeComponent();
            blazorWebView.BlazorWebViewInitializing += (_, e) =>
                e.UserDataFolder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "JBC.ExploreTheWorld", "WebView2");
            // Use absolute path so BlazorWebView resolves correctly in COM-hosted context
            // where AppContext.BaseDirectory may point to the Office application directory.
            blazorWebView.HostPage  = Path.Combine(
                Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)!,
                "wwwroot", "index.html");
            blazorWebView.Services  = serviceProvider;
            blazorWebView.RootComponents.Add(
                new RootComponent("#app", typeof(Routes), null));
            blazorWebView.StartPath = "/countries-now";
        }
    }
}
