using System;
using System.Collections.Generic;
using System.Globalization;
using JBC.ExploreTheWorld.DL.MsOfficeApi.MsOffice;
using PP = JBC.ExploreTheWorld.DL.MsOfficeApi.MsPowerPoint;
using Undefined = JBC.ExploreTheWorld.DL.MsOfficeApi.MsOffice.MsOfficeUndefined;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.PowerPoint
{
    // Maps the Office.js-shaped rows to the canonical JBC.ExploreTheWorld.DL.MsOfficeApi.MsPowerPoint
    // entities and back. VBA properties the Office.js object model cannot provide are set
    // to "**Undefined" (strings), -99 (enums/numerics) or null (booleans/whole objects).
    public static class MsPowerPointJsMapper
    {
        public static PP.Presentation ToPresentation(PowerPointPresentationJs_Row source, string? fileName)
        {
            var presentation = new PP.Presentation
            {
                SlideMaster = null,
                TitleMaster = null,
                HandoutMaster = null,
                NotesMaster = null,
                Coauthoring = null,
                PageSetup = null,
                PrintOptions = null,
                SlideShowSettings = null,
                BuiltInDocumentProperties = new List<PP.BuiltInDocumentProperty>(),
                CustomDocumentProperties = new List<PP.CustomDocumentProperty>(),
                ContentTypeProperties = null,
                ColorSchemes = null,
                CustomXmlParts = null,
                Designs = null,
                ExtraColors = null,
                Fonts = null,
                Guides = null,
                SectionProperties = new List<PP.SectionProperty>(),
                Slides = new List<PP.Slide>(),
                Tags = new List<PP.Tag>(),
                Name = string.IsNullOrEmpty(fileName) ? Undefined.String : fileName,
                Path = Undefined.String,
                FullName = Undefined.String,
                AutoSaveOn = null,
                ChartDataPointTrack = null,
                CreateVideoStatus = Undefined.Number,
                DefaultLanguageID = Undefined.Number,
                DisplayComments = Undefined.Number,
                EncryptionProvider = Undefined.String,
                EnvelopeVisible = Undefined.Number,
                FarEastLineBreakLanguage = Undefined.Number,
                FarEastLineBreakLevel = Undefined.Number,
                Final = null,
                GridDistance = Undefined.String,
                HasHandoutMaster = Undefined.Number,
                HasNotesMaster = Undefined.Number,
                HasTitleMaster = Undefined.Number,
                HasVBProject = Undefined.Number,
                InMergeMode = null,
                LayoutDirection = Undefined.Number,
                NoLineBreakAfter = Undefined.String,
                NoLineBreakBefore = Undefined.String,
                Password = Undefined.String,
                PasswordEncryptionAlgorithm = Undefined.String,
                PasswordEncryptionFileProperties = null,
                PasswordEncryptionKeyLength = Undefined.NumberLong,
                PasswordEncryptionProvider = Undefined.String,
                ReadOnly = Undefined.Number,
                ReadOnlyRecommended = null,
                RemovePersonalInformation = Undefined.Number,
                Saved = Undefined.Number,
                SnapToGrid = Undefined.Number,
                TemplateName = Undefined.String,
                VBASigned = Undefined.Number,
                WritePassword = Undefined.String,
            };

            foreach (var slide in source.Slides)
            {
                presentation.Slides.Add(ToSlide(slide));
            }

            return presentation;
        }

        private static PP.Slide ToSlide(PowerPointSlideJs_Row source)
        {
            var slide = new PP.Slide
            {
                SlideShowTransition = null,
                ColorScheme = null,
                ThemeColorScheme = null,
                HeadersFooters = null,
                NotesPage = null,
                Background = null,
                Shapes = new List<PP.Shape>(),
                Tags = new List<PP.Tag>(),
                Comments = null,
                Hyperlinks = null,
                Name = Undefined.String,
                CustomLayout_Name = source.LayoutName ?? Undefined.String,
                Design_Name = Undefined.String,
                SlideID = Undefined.Number,
                SlideIndex = source.Index,
                SlideNumber = source.Index,
                sectionIndex = Undefined.Number,
                BackgroundStyle = Undefined.Number,
                DisplayMasterShapes = Undefined.Number,
                FollowMasterBackground = Undefined.Number,
                HasNotesPage = Undefined.Number,
                Layout = Undefined.Number,
                PrintSteps = Undefined.NumberLong,
            };

            foreach (var shape in source.Shapes)
            {
                slide.Shapes.Add(ToShape(shape));
            }

            return slide;
        }

        private static PP.Shape ToShape(PowerPointShapeJs_Row source)
        {
            bool hasText = !string.IsNullOrEmpty(source.Text);

            var shape = new PP.Shape
            {
                AnimationSettings = null,
                Fill = null,
                Line = null,
                PictureFormat = null,
                PlaceholderFormat = null,
                TextFrame = hasText ? new PP.TextFrame { TextRange = ToTextRange(source.Text!) } : null,
                TextFrame2 = hasText ? new PP.TextFrame2 { TextRange2 = ToTextRange2(source.Text!) } : null,
                Image = null,
                ActionSettings = null,
                Tags = new List<PP.Tag>(),
                GroupItems = null,
                Id = Undefined.NumberLong,
                Name = source.Name ?? string.Empty,
                Type = MsoShapeTypeFromJs(source.Type),
                AutoShapeType = Undefined.Number,
                AlternativeText = source.AltTextDescription ?? string.Empty,
                Title = source.AltTextTitle ?? string.Empty,
                Left = PointsString(source.Left),
                Top = PointsString(source.Top),
                Width = PointsString(source.Width),
                Height = PointsString(source.Height),
                Rotation = Undefined.String,
                ZOrderPosition = source.ZOrderIndex,
                Visible = -1,
                BackgroundStyle = Undefined.Number,
                BlackWhiteMode = Undefined.Number,
                Child = Undefined.Number,
                ConnectionSiteCount = Undefined.NumberLong,
                Connector = Undefined.Number,
                Creator = Undefined.NumberLong,
                Decorative = Undefined.Number,
                GraphicStyle = Undefined.Number,
                HasChart = Undefined.Number,
                HasInkXML = Undefined.Number,
                HasSectionZoom = Undefined.Number,
                HasSmartArt = Undefined.Number,
                HasTable = Undefined.Number,
                HasTextFrame = hasText ? -1 : 0,
                HorizontalFlip = Undefined.Number,
                InkXML = Undefined.String,
                IsNarration = Undefined.Number,
                LockAspectRatio = Undefined.Number,
                ShapeStyle = Undefined.Number,
                VerticalFlip = Undefined.Number,
                Vertices = Undefined.String,
                MediaType = Undefined.Number,
            };

            return shape;
        }

        private static PP.TextRange ToTextRange(string text)
        {
            return new PP.TextRange
            {
                Font = null,
                Length = text.Length,
                Start = 1,
                Text = text,
            };
        }

        private static PP.TextRange2 ToTextRange2(string text)
        {
            return new PP.TextRange2
            {
                Font = null,
                Length = text.Length,
                Start = 1,
                Text = text,
            };
        }

        // Office.js PowerPoint.ShapeType -> VBA MsoShapeType.
        private static int MsoShapeTypeFromJs(string? shapeType)
        {
            return shapeType?.ToLowerInvariant() switch
            {
                "geometricshape" => 1, // msoAutoShape
                "callout" => 2, // msoCallout
                "chart" => 3, // msoChart
                "freeform" => 5, // msoFreeform
                "group" => 6, // msoGroup
                "line" => 9, // msoLine
                "image" => 13, // msoPicture
                "placeholder" => 14, // msoPlaceholder
                "media" => 16, // msoMedia
                "textbox" => 17, // msoTextBox
                "table" => 19, // msoTable
                "smartart" => 24, // msoSmartArt
                "contentapp" => 27, // msoContentApp
                _ => Undefined.Number,
            };
        }

        private static string PointsString(double? points)
        {
            if (!points.HasValue) return Undefined.String;

            return ((float)points.Value).ToString(CultureInfo.InvariantCulture);
        }

        public static PowerPointPresentationJs_Row FromPresentation(PP.Presentation presentation)
        {
            var row = new PowerPointPresentationJs_Row
            {
                Title = presentation.Name,
                Slides = new List<PowerPointSlideJs_Row>(),
            };

            foreach (var slide in presentation.Slides ?? new List<PP.Slide>())
            {
                var slideRow = new PowerPointSlideJs_Row
                {
                    Id = slide.SlideID?.ToString(CultureInfo.InvariantCulture),
                    Index = slide.SlideIndex ?? 0,
                    LayoutName = slide.CustomLayout_Name,
                    Shapes = new List<PowerPointShapeJs_Row>(),
                };

                int zOrderIndex = 0;
                foreach (var shape in slide.Shapes ?? new List<PP.Shape>())
                {
                    zOrderIndex++;
                    slideRow.Shapes.Add(new PowerPointShapeJs_Row
                    {
                        Id = shape.Id?.ToString(CultureInfo.InvariantCulture),
                        Name = shape.Name,
                        Type = JsShapeTypeFromMso(shape.Type),
                        Left = PointsFromString(shape.Left),
                        Top = PointsFromString(shape.Top),
                        Width = PointsFromString(shape.Width),
                        Height = PointsFromString(shape.Height),
                        AltTextDescription = shape.AlternativeText,
                        AltTextTitle = shape.Title,
                        ZOrderIndex = shape.ZOrderPosition ?? zOrderIndex,
                        Text = shape.TextFrame?.TextRange?.Text,
                    });
                }

                row.Slides.Add(slideRow);
            }

            return row;
        }

        private static string JsShapeTypeFromMso(int? msoShapeType)
        {
            return msoShapeType switch
            {
                1 => "GeometricShape",
                2 => "Callout",
                3 => "Chart",
                5 => "Freeform",
                6 => "Group",
                9 => "Line",
                13 => "Image",
                14 => "Placeholder",
                16 => "Media",
                17 => "TextBox",
                19 => "Table",
                24 => "SmartArt",
                27 => "ContentApp",
                _ => "Unsupported",
            };
        }

        private static double? PointsFromString(string? points)
        {
            if (string.IsNullOrEmpty(points) || MsOfficeUndefined.IsUndefined(points)) return null;

            return double.TryParse(points, NumberStyles.Float, CultureInfo.InvariantCulture, out double value)
                ? value
                : null;
        }
    }
}
