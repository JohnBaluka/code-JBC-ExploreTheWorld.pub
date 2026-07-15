using Microsoft.AspNetCore.Components.Web;
using Microsoft.AspNetCore.Components;

namespace JBC.ExploreTheWorld.AL.BlazorLib
{
    public class RenderMode_AppService
    {
        private IComponentRenderMode? _renderMode = null; // null = static SSR (default)

        public void SetRenderMode(IComponentRenderMode renderMode)
        {
            _renderMode = renderMode;
        }

        public IComponentRenderMode? GetRenderMode()
        {
            return _renderMode;
        }
    }
}
