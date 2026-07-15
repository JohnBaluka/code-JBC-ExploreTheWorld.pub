using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;
using Radzen;

namespace JBC.ExploreTheWorld.AL.BlazorLib
{
    /// <summary>
    /// Base class for Radzen components that need a Hidden parameter.
    /// Inherits RadzenComponent and adds a Hidden parameter.
    /// Provides standard Radzen service injections.
    /// When Hidden is true (and Visible is true), the component renders with display: none;
    /// so it remains in the DOM but is visually hidden.
    /// When Visible is false, the component does not render at all.
    /// </summary>
    public class Base__RadzenComponent : RadzenComponent
    {
        [Inject] protected IJSRuntime JS { get; set; } = default!;
        [Inject] protected NavigationManager NavigationManager { get; set; } = default!;
        [Inject] protected DialogService Dialog_AppService { get; set; } = default!;
        [Inject] protected TooltipService Tooltip_AppService { get; set; } = default!;
        [Inject] protected ContextMenuService ContextMenu_AppService { get; set; } = default!;
        [Inject] protected NotificationService Notification_AppService { get; set; } = default!;
        /// <summary>
        /// When true, adds "display: none;" to the root element's style.
        /// The component still renders in the DOM (unlike Visible=false which prevents rendering entirely).
        /// Defaults to false.
        /// </summary>
        [Parameter]
        public bool Hidden { get; set; } = false;

        /// <summary>
        /// Returns the combined style string, prepending "display: none;" when Hidden is true.
        /// Use this in the root element's style attribute instead of @Style.
        /// </summary>
        protected string GetHiddenStyle()
        {
            if (Hidden)
            {
                return string.IsNullOrEmpty(Style)
                    ? "display: none;"
                    : $"display: none; {Style}";
            }

            return Style ?? "";
        }

        public bool DebugRadzen { get; set; } = false;
        public string DebugMessage { get; set; } = "";
    }
}
