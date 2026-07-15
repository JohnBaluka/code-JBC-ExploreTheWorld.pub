using JBC.ExploreTheWorld.CL;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json.Serialization;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsPowerPoint
{
    [Serializable]
    public partial class Slide : Slide_Fields
    {
        // Entities
        public SlideShowTransition? SlideShowTransition { get; set; }
        public ColorScheme? ColorScheme { get; set; }
        public ThemeColorScheme? ThemeColorScheme { get; set; }
        public HeadersFooters? HeadersFooters { get; set; }
        public SlideRange? NotesPage { get; set; }

        // Lists
        public List<Shape>? Background { get; set; }
        public List<Shape>? Shapes { get; set; }
        public List<Tag>? Tags { get; set; }
        public List<Comment>? Comments { get; set; }
        public List<Hyperlink>? Hyperlinks { get; set; }

        [JsonIgnore]
        public string? SVG { get; set; }

        public string GetNotes()
        {
            if (NotesPage == null) return string.Empty;
            if (NotesPage.Shapes == null) return string.Empty;

            var oNotesPageShape = NotesPage.Shapes.Where(o => o.Name != null && o.Name.Contains("Notes Placeholder")).FirstOrDefault();

            if (oNotesPageShape == null) return string.Empty;
            if (oNotesPageShape.TextFrame == null) return string.Empty;
            if (oNotesPageShape.TextFrame.TextRange == null) return string.Empty;

            string text = oNotesPageShape.TextFrame.TextRange.Text ?? string.Empty;

            string noBreakSpace = "\u00A0";
            text = text.Replace(noBreakSpace, "");

            return text;
        }

        public string GetTitle()
        {
            if (Shapes == null) return string.Empty;

            var oShape = Shapes.Where(o => o.Name != null && o.Name.Contains("Title")).FirstOrDefault();

            if (oShape == null) return string.Empty;
            if (oShape.TextFrame == null) return string.Empty;
            if (oShape.TextFrame.TextRange == null) return string.Empty;

            return oShape.TextFrame.TextRange.Text ?? string.Empty;
        }

        public List<Shape> GetShapesWithAnimations(bool includeNarrations = false)
        {
            if (Shapes == null) return new List<Shape>();

            if (includeNarrations)
                return Shapes.Where(o => o.AnimationSettings != null && o.AnimationSettings.AnimationOrder.HasValue).OrderBy(o => o.AnimationSettings!.AnimationOrder).ToList();

            return Shapes.Where(o => o.AnimationSettings != null && o.AnimationSettings.AnimationOrder.HasValue && o.IsNarrationAsBoolean() == false).OrderBy(o => o.AnimationSettings!.AnimationOrder).ToList();
        }
    }
}
