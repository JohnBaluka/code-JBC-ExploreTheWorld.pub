using System;
using JBC.ExploreTheWorld.AL.BlazorLib;

namespace JBC.ExploreTheWorld.AL.WinFormsLib
{
    public class WinFormNewWindow_AppService : NewWindow_AppService__Interface
    {
        private readonly IServiceProvider _serviceProvider;

        public WinFormNewWindow_AppService(IServiceProvider serviceProvider)
        {
            _serviceProvider = serviceProvider;
        }

        public void OpenNewWindow()
        {
            // Must run on the STA thread; use the first open form as dispatcher.
            if (System.Windows.Forms.Application.OpenForms.Count > 0)
                System.Windows.Forms.Application.OpenForms[0]!.BeginInvoke(() =>
                    new ExploreTheWorld_Form(_serviceProvider).Show());
        }
    }
}
