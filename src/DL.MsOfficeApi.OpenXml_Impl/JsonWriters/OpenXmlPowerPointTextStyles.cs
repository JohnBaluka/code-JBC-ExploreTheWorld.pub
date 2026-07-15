using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using D = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.OpenXml_Impl.JsonWriters
{
    // Builds the ordered list of run-property sources a text run inherits from:
    // run rPr -> paragraph defRPr -> shape lstStyle -> placeholder chain lstStyles ->
    // master text styles -> presentation defaultTextStyle.
    internal static class OpenXmlPowerPointTextStyles
    {
        // One inheritance chain for a run at a given outline level, split at the point
        // where the shape's style fontRef color slots in (near = run/paragraph/shape
        // sources, far = placeholder chain and master/presentation styles).
        public static (List<OpenXmlElement> Near, List<OpenXmlElement> Far) BuildChain(OpenXmlElement? runProperties,
            D.Paragraph paragraph, P.Shape shape, OpenXmlPowerPointContext context)
        {
            var near = new List<OpenXmlElement>();
            var far = new List<OpenXmlElement>();
            int level = (paragraph.ParagraphProperties?.Level?.Value ?? 0) + 1;

            if (runProperties != null) near.Add(runProperties);

            if (paragraph.ParagraphProperties?.GetFirstChild<D.DefaultRunProperties>() is D.DefaultRunProperties paragraphDefaults)
            {
                near.Add(paragraphDefaults);
            }

            AddLevelProperties(near, shape.TextBody?.ListStyle, level);

            var placeholder = OpenXmlPowerPointContext.PlaceholderOf(shape);
            if (placeholder != null)
            {
                foreach (var inherited in context.GetPlaceholderChain(placeholder))
                {
                    AddLevelProperties(far, inherited.TextBody?.ListStyle, level);
                }
            }

            if (placeholder != null || !IsTextBox(shape))
            {
                AddLevelProperties(far, context.GetMasterTextStyle(placeholder), level);
            }

            AddLevelProperties(far, context.PresentationPart.Presentation?.DefaultTextStyle, level);

            return (near, far);
        }

        public static bool IsTextBox(P.Shape shape)
        {
            return shape.NonVisualShapeProperties?.NonVisualShapeDrawingProperties?.TextBox?.Value == true;
        }

        private static void AddLevelProperties(List<OpenXmlElement> chain, OpenXmlCompositeElement? listStyle, int level)
        {
            if (listStyle == null) return;

            OpenXmlElement? levelProperties = level switch
            {
                1 => listStyle.GetFirstChild<D.Level1ParagraphProperties>(),
                2 => listStyle.GetFirstChild<D.Level2ParagraphProperties>(),
                3 => listStyle.GetFirstChild<D.Level3ParagraphProperties>(),
                4 => listStyle.GetFirstChild<D.Level4ParagraphProperties>(),
                5 => listStyle.GetFirstChild<D.Level5ParagraphProperties>(),
                6 => listStyle.GetFirstChild<D.Level6ParagraphProperties>(),
                7 => listStyle.GetFirstChild<D.Level7ParagraphProperties>(),
                8 => listStyle.GetFirstChild<D.Level8ParagraphProperties>(),
                9 => listStyle.GetFirstChild<D.Level9ParagraphProperties>(),
                _ => null,
            };

            // p:txStyles/p:notesStyle wrap the lvlNpPr in a nested a:lstStyle-shaped
            // element; a:lstStyle holds them directly. Level 1 also falls back to defPPr.
            if (levelProperties == null && level == 1)
            {
                levelProperties = listStyle.GetFirstChild<D.DefaultParagraphProperties>();
            }

            if (levelProperties?.GetFirstChild<D.DefaultRunProperties>() is D.DefaultRunProperties defaults)
            {
                chain.Add(defaults);
            }
        }

        // First non-null value along the chain.
        public static T? Resolve<T>(List<OpenXmlElement> chain, Func<OpenXmlElement, T?> selector)
            where T : class
        {
            foreach (var source in chain)
            {
                if (selector(source) is T value) return value;
            }

            return null;
        }

        public static T? ResolveValue<T>(List<OpenXmlElement> chain, Func<OpenXmlElement, T?> selector)
            where T : struct
        {
            foreach (var source in chain)
            {
                if (selector(source) is T value) return value;
            }

            return null;
        }
    }
}
