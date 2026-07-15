using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using JBC.ExploreTheWorld.DL.MsOfficeApi.MsOffice;
using NetOffice.OfficeApi.Enums;
using NetOffice.PowerPointApi;
using NetOffice.PowerPointApi.Enums;
using Office = NetOffice.OfficeApi;
using PP = JBC.ExploreTheWorld.DL.MsOfficeApi.MsPowerPoint;
using static JBC.ExploreTheWorld.DL.MsOfficeApi.NetOffice_Impl.JsonWriters.ComTryGet;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.NetOffice_Impl.JsonWriters
{
    // Builds the canonical JBC.ExploreTheWorld.DL.MsOfficeApi.MsPowerPoint entity graph from a running
    // PowerPoint presentation (via the strongly-typed NetOffice object model) and writes
    // it with MsOfficeJsonSerializer. The output matches MSO_MsPowerPoint_JsonWriter.bas.
    public static class MsPowerPointJsonWriter
    {
        public static void WritePresentationToJsonFile(Presentation presentation, string outputFilePath,
            MsOfficeJsonWriterOptions? options = null, Action<string>? log = null)
        {
            options ??= new MsOfficeJsonWriterOptions();

            log?.Invoke($"Writing JSON: {outputFilePath}");

            PP.Presentation entity = BuildPresentation(presentation, outputFilePath, options, log);

            MsOfficeJsonSerializer.WriteToFile(entity, outputFilePath);

            log?.Invoke($"Done: {outputFilePath}");
        }

        private static PP.Presentation BuildPresentation(Presentation p, string outputFilePath, MsOfficeJsonWriterOptions options, Action<string>? log)
        {
            var prs = new PP.Presentation();

            // Entities
            prs.SlideMaster = Obj(() => BuildMaster(p.SlideMaster, true, outputFilePath, options));
            prs.TitleMaster = Obj<PP.SlideMaster>(() =>
            {
                if (p.HasTitleMaster != MsoTriState.msoTrue) return null;
                return BuildMaster(p.TitleMaster, false, outputFilePath, options);
            });
            prs.HandoutMaster = Obj(() => BuildMaster(p.HandoutMaster, false, outputFilePath, options));
            prs.NotesMaster = Obj(() => BuildMaster(p.NotesMaster, false, outputFilePath, options));
            prs.Coauthoring = Obj(() => new PP.Coauthoring
            {
                FavorServerEditsDuringMerge = p.Coauthoring.FavorServerEditsDuringMerge,
                MergeMode = p.Coauthoring.MergeMode,
                PendingUpdates = p.Coauthoring.PendingUpdates,
            });
            prs.PageSetup = Obj(() => BuildPageSetup(p.PageSetup));
            prs.PrintOptions = Obj(() => BuildPrintOptions(p.PrintOptions));
            prs.SlideShowSettings = Obj(() => BuildSlideShowSettings(p.SlideShowSettings));

            // Lists
            prs.BuiltInDocumentProperties = Obj(() => BuildDocumentProperties(p.BuiltInDocumentProperties)) ?? new List<PP.BuiltInDocumentProperty>();
            prs.CustomDocumentProperties = Obj(() => BuildCustomDocumentProperties(p.CustomDocumentProperties)) ?? new List<PP.CustomDocumentProperty>();
            prs.ContentTypeProperties = Obj(() => BuildContentTypeProperties(p.ContentTypeProperties));
            prs.ColorSchemes = Obj<List<PP.ColorScheme>>(() =>
            {
                var list = new List<PP.ColorScheme>();
                foreach (ColorScheme scheme in p.ColorSchemes) list.Add(BuildColorScheme(scheme));
                return list;
            });
            prs.CustomXmlParts = Obj<List<PP.CustomXmlPart>>(() =>
            {
                var list = new List<PP.CustomXmlPart>();
                foreach (Office.CustomXMLPart part in p.CustomXMLParts)
                {
                    list.Add(new PP.CustomXmlPart
                    {
                        BuiltIn = Bool(() => part.BuiltIn),
                        Id = Str(() => part.Id),
                        NamespaceURI = Str(() => part.NamespaceURI),
                        XML = Str(() => part.XML),
                    });
                }
                return list;
            });
            prs.Designs = Obj<List<PP.Design>>(() =>
            {
                var list = new List<PP.Design>();
                foreach (Design design in p.Designs)
                {
                    list.Add(new PP.Design
                    {
                        Name = Str(() => design.Name),
                        Index = Long(() => design.Index),
                        HasTitleMaster = Int(() => design.HasTitleMaster),
                        Preserved = Int(() => design.Preserved),
                    });
                }
                return list;
            });
            prs.ExtraColors = Obj<List<long>>(() =>
            {
                var list = new List<long>();
                int count = p.ExtraColors.Count;
                for (int i = 1; i <= count; i++) list.Add(p.ExtraColors[i]);
                return list;
            });
            prs.Fonts = Obj<List<PP.PresentationFont>>(() =>
            {
                var list = new List<PP.PresentationFont>();
                foreach (Font font in p.Fonts)
                {
                    list.Add(new PP.PresentationFont
                    {
                        Name = Str(() => font.Name),
                        Embeddable = Int(() => font.Embeddable),
                        Embedded = Int(() => font.Embedded),
                    });
                }
                return list;
            });
            prs.Guides = Obj<List<PP.Guide>>(() =>
            {
                var list = new List<PP.Guide>();
                int count = p.Guides.Count;
                for (int i = 1; i <= count; i++)
                {
                    Guide guide = p.Guides[i];
                    list.Add(new PP.Guide
                    {
                        Color = Obj(() => BuildColorFormat(guide.Color)),
                        Orientation = Int(() => guide.Orientation),
                        Position = SingleStr(() => guide.Position),
                    });
                }
                return list;
            });
            prs.SectionProperties = Obj<List<PP.SectionProperty>>(() =>
            {
                var list = new List<PP.SectionProperty>();
                int count = p.SectionProperties.Count;
                for (int i = 1; i <= count; i++)
                {
                    int index = i;
                    list.Add(new PP.SectionProperty
                    {
                        Index = index,
                        Name = Str(() => p.SectionProperties.Name(index)),
                        SectionID = Str(() => p.SectionProperties.SectionID(index)),
                        SlidesCount = Int(() => p.SectionProperties.SlidesCount(index)),
                    });
                }
                return list;
            }) ?? new List<PP.SectionProperty>();
            prs.Slides = Obj<List<PP.Slide>>(() =>
            {
                var list = new List<PP.Slide>();
                int count = p.Slides.Count;
                log?.Invoke($"  Writing {count} slides...");
                int index = 0;
                foreach (Slide slide in p.Slides)
                {
                    index++;
                    if (index % 10 == 0) log?.Invoke($"  Slide {index}/{count}...");
                    list.Add(BuildSlide(slide, outputFilePath, options));
                }
                return list;
            }) ?? new List<PP.Slide>();
            prs.Tags = Obj(() => BuildTags(p.Tags)) ?? new List<PP.Tag>();

            // Fields
            prs.Name = Str(() => p.Name);
            prs.Path = Str(() => p.Path);
            prs.FullName = Str(() => p.FullName);
            prs.AutoSaveOn = Bool(() => p.AutoSaveOn);
            prs.ChartDataPointTrack = Bool(() => p.ChartDataPointTrack);
            prs.CreateVideoStatus = Int(() => p.CreateVideoStatus);
            prs.DefaultLanguageID = Int(() => p.DefaultLanguageID);
            prs.DisplayComments = Int(() => p.DisplayComments);
            prs.EncryptionProvider = Str(() => p.EncryptionProvider);
            prs.EnvelopeVisible = Int(() => p.EnvelopeVisible);
            prs.FarEastLineBreakLanguage = Int(() => p.FarEastLineBreakLanguage);
            prs.FarEastLineBreakLevel = Int(() => p.FarEastLineBreakLevel);
            prs.Final = Bool(() => p.Final);
            prs.GridDistance = SingleStr(() => p.GridDistance);
            // NetOffice surfaces these MsoTriState properties as bool; the canonical
            // schema stores the MsoTriState value (msoTrue = -1), like the VBA writer.
            prs.HasHandoutMaster = Int(() => p.HasHandoutMaster ? (int)MsoTriState.msoTrue : (int)MsoTriState.msoFalse);
            prs.HasNotesMaster = Int(() => p.HasNotesMaster ? (int)MsoTriState.msoTrue : (int)MsoTriState.msoFalse);
            prs.HasTitleMaster = Int(() => p.HasTitleMaster);
            prs.HasVBProject = Int(() => p.HasVBProject ? (int)MsoTriState.msoTrue : (int)MsoTriState.msoFalse);
            prs.InMergeMode = Bool(() => p.InMergeMode);
            prs.LayoutDirection = Int(() => p.LayoutDirection);
            prs.NoLineBreakAfter = Str(() => p.NoLineBreakAfter);
            prs.NoLineBreakBefore = Str(() => p.NoLineBreakBefore);
            prs.Password = Str(() => p.Password);
            prs.PasswordEncryptionAlgorithm = Str(() => p.PasswordEncryptionAlgorithm);
            prs.PasswordEncryptionFileProperties = Bool(() => p.PasswordEncryptionFileProperties);
            prs.PasswordEncryptionKeyLength = Long(() => p.PasswordEncryptionKeyLength);
            prs.PasswordEncryptionProvider = Str(() => p.PasswordEncryptionProvider);
            prs.ReadOnly = Int(() => p.ReadOnly);
            prs.ReadOnlyRecommended = Bool(() => p.ReadOnlyRecommended);
            prs.RemovePersonalInformation = Int(() => p.RemovePersonalInformation);
            prs.Saved = Int(() => p.Saved);
            prs.SnapToGrid = Int(() => p.SnapToGrid);
            prs.TemplateName = Str(() => p.TemplateName);
            prs.VBASigned = Int(() => p.VBASigned);
            prs.WritePassword = Str(() => p.WritePassword);

            return prs;
        }

        // -- Masters ------------------------------------------------------------

        private static PP.SlideMaster BuildMaster(_Master master, bool isSlideMaster, string outputFilePath, MsOfficeJsonWriterOptions options)
        {
            var entity = new PP.SlideMaster
            {
                SlideShowTransition = Obj(() => BuildSlideShowTransition(master.SlideShowTransition)),
                SlideMasterLayouts = isSlideMaster
                    ? Obj<List<PP.SlideMasterLayout>>(() =>
                    {
                        var list = new List<PP.SlideMasterLayout>();
                        foreach (CustomLayout layout in master.CustomLayouts) list.Add(BuildSlideMasterLayout(layout, outputFilePath, options));
                        return list;
                    })
                    : null,
                Shapes = BuildShapes(master.Shapes, outputFilePath, options),
                Name = Str(() => master.Name),
            };

            return entity;
        }

        private static PP.SlideMasterLayout BuildSlideMasterLayout(CustomLayout layout, string outputFilePath, MsOfficeJsonWriterOptions options)
        {
            return new PP.SlideMasterLayout
            {
                SlideShowTransition = Obj(() => BuildSlideShowTransition(layout.SlideShowTransition)),
                Shapes = BuildShapes(layout.Shapes, outputFilePath, options),
                Hyperlinks = Obj<List<PP.Hyperlink>>(() =>
                {
                    int count = layout.Hyperlinks.Count;
                    if (count == 0) return null;
                    return BuildHyperlinks(layout.Hyperlinks);
                }),
                Name = Str(() => layout.Name),
                Index = Int(() => layout.Index),
            };
        }

        // -- Presentation sub-objects ---------------------------------------------

        private static PP.PageSetup BuildPageSetup(PageSetup pageSetup)
        {
            return new PP.PageSetup
            {
                FirstSlideNumber = Long(() => pageSetup.FirstSlideNumber),
                NotesOrientation = Int(() => pageSetup.NotesOrientation),
                SlideHeight = SingleStr(() => pageSetup.SlideHeight),
                SlideOrientation = Int(() => pageSetup.SlideOrientation),
                SlideSize = Int(() => pageSetup.SlideSize),
                SlideWidth = SingleStr(() => pageSetup.SlideWidth),
            };
        }

        private static PP.PrintOptions BuildPrintOptions(PrintOptions printOptions)
        {
            return new PP.PrintOptions
            {
                ActivePrinter = Str(() => printOptions.ActivePrinter),
                Collate = Int(() => printOptions.Collate),
                FitToPage = Int(() => printOptions.FitToPage),
                FrameSlides = Int(() => printOptions.FrameSlides),
                HandoutOrder = Int(() => printOptions.HandoutOrder),
                HighQuality = Int(() => printOptions.HighQuality),
                NumberOfCopies = Long(() => printOptions.NumberOfCopies),
                OutputType = Int(() => printOptions.OutputType),
                PrintColorType = Int(() => printOptions.PrintColorType),
                PrintComments = Int(() => printOptions.PrintComments),
                PrintFontsAsGraphics = Int(() => printOptions.PrintFontsAsGraphics),
                PrintHiddenSlides = Int(() => printOptions.PrintHiddenSlides),
                PrintInBackground = Int(() => printOptions.PrintInBackground),
                RangeType = Int(() => printOptions.RangeType),
                SlideShowName = Str(() => printOptions.SlideShowName),
            };
        }

        private static PP.SlideShowSettings BuildSlideShowSettings(SlideShowSettings settings)
        {
            return new PP.SlideShowSettings
            {
                PointerColor = Obj(() => BuildColorFormat(settings.PointerColor)),
                AdvanceMode = Int(() => settings.AdvanceMode),
                EndingSlide = Long(() => settings.EndingSlide),
                LoopUntilStopped = Int(() => settings.LoopUntilStopped),
                RangeType = Int(() => settings.RangeType),
                ShowMediaControls = Int(() => settings.ShowMediaControls),
                ShowPresenterView = Int(() => settings.ShowPresenterView),
                ShowScrollbar = Int(() => settings.ShowScrollbar),
                ShowType = Int(() => settings.ShowType),
                ShowWithAnimation = Int(() => settings.ShowWithAnimation),
                ShowWithNarration = Int(() => settings.ShowWithNarration),
                StartingSlide = Long(() => settings.StartingSlide),
            };
        }

        private static List<PP.BuiltInDocumentProperty> BuildDocumentProperties(object properties)
        {
            var list = new List<PP.BuiltInDocumentProperty>();

            foreach (Office.DocumentProperty property in (Office.DocumentProperties)properties)
            {
                list.Add(new PP.BuiltInDocumentProperty
                {
                    Name = Str(() => property.Name),
                    Creator = Long(() => property.Creator),
                    LinkSource = Str(() => property.LinkSource),
                    LinkToContent = Int(() => property.LinkToContent),
                    Type = Int(() => property.Type),
                    Value = Str(() => property.Value),
                });
            }

            return list;
        }

        private static List<PP.CustomDocumentProperty> BuildCustomDocumentProperties(object properties)
        {
            var list = new List<PP.CustomDocumentProperty>();

            foreach (Office.DocumentProperty property in (Office.DocumentProperties)properties)
            {
                list.Add(new PP.CustomDocumentProperty
                {
                    Name = Str(() => property.Name),
                    Creator = Long(() => property.Creator),
                    LinkSource = Str(() => property.LinkSource),
                    LinkToContent = Int(() => property.LinkToContent),
                    Type = Int(() => property.Type),
                    Value = Str(() => property.Value),
                });
            }

            return list;
        }

        private static List<PP.ContentTypeProperty> BuildContentTypeProperties(Office.MetaProperties properties)
        {
            var list = new List<PP.ContentTypeProperty>();

            foreach (Office.MetaProperty property in properties)
            {
                list.Add(new PP.ContentTypeProperty
                {
                    Id = Str(() => property.Id),
                    Name = Str(() => property.Name),
                    Creator = Long(() => property.Creator),
                    IsReadOnly = Bool(() => property.IsReadOnly),
                    IsRequired = Bool(() => property.IsRequired),
                    Type = Int(() => property.Type),
                    Value = Str(() => property.Value),
                });
            }

            return list;
        }

        // -- Slides ---------------------------------------------------------------

        private static PP.Slide BuildSlide(Slide slide, string outputFilePath, MsOfficeJsonWriterOptions options)
        {
            return new PP.Slide
            {
                SlideShowTransition = Obj(() => BuildSlideShowTransition(slide.SlideShowTransition)),
                ColorScheme = Obj(() => BuildColorScheme(slide.ColorScheme)),
                ThemeColorScheme = Obj(() => BuildThemeColorScheme(slide.ThemeColorScheme)),
                HeadersFooters = Obj(() => BuildHeadersFooters(slide.HeadersFooters)),
                NotesPage = Obj(() => new PP.SlideRange
                {
                    Shapes = BuildShapes(slide.NotesPage.Shapes, outputFilePath, options),
                    Name = Str(() => slide.NotesPage.Name),
                }),
                Background = Obj<List<PP.Shape>>(() =>
                {
                    var list = new List<PP.Shape>();
                    foreach (Shape shape in slide.Background) list.Add(BuildShape(shape, outputFilePath, options));
                    return list;
                }),
                Shapes = BuildShapes(slide.Shapes, outputFilePath, options),
                Tags = Obj(() => BuildTags(slide.Tags)),
                Comments = Obj<List<PP.Comment>>(() =>
                {
                    int count = slide.Comments.Count;
                    if (count == 0) return null;
                    var list = new List<PP.Comment>();
                    for (int i = 1; i <= count; i++) list.Add(BuildComment(slide.Comments[i]));
                    return list;
                }),
                Hyperlinks = Obj<List<PP.Hyperlink>>(() =>
                {
                    int count = slide.Hyperlinks.Count;
                    if (count == 0) return null;
                    return BuildHyperlinks(slide.Hyperlinks);
                }),
                Name = Str(() => slide.Name),
                CustomLayout_Name = Str(() => slide.CustomLayout.Name),
                Design_Name = Str(() => slide.Design.Name),
                SlideID = Int(() => slide.SlideID),
                SlideIndex = Int(() => slide.SlideIndex),
                SlideNumber = Int(() => slide.SlideNumber),
                sectionIndex = Int(() => slide.sectionIndex),
                BackgroundStyle = Int(() => slide.BackgroundStyle),
                DisplayMasterShapes = Int(() => slide.DisplayMasterShapes),
                FollowMasterBackground = Int(() => slide.FollowMasterBackground),
                HasNotesPage = Int(() => slide.HasNotesPage),
                Layout = Int(() => slide.Layout),
                PrintSteps = Long(() => slide.PrintSteps),
            };
        }

        private static PP.HeadersFooters BuildHeadersFooters(HeadersFooters headersFooters)
        {
            return new PP.HeadersFooters
            {
                DateAndTime = Obj(() => BuildHeaderFooter(headersFooters.DateAndTime)),
                Footer = Obj(() => BuildHeaderFooter(headersFooters.Footer)),
                Header = Obj(() => BuildHeaderFooter(headersFooters.Header)),
                SlideNumber = Obj(() => BuildHeaderFooter(headersFooters.SlideNumber)),
                DisplayOnTitleSlide = Int(() => headersFooters.DisplayOnTitleSlide),
            };
        }

        private static PP.HeaderFooter BuildHeaderFooter(HeaderFooter headerFooter)
        {
            return new PP.HeaderFooter
            {
                Format = Int(() => headerFooter.Format),
                Text = Str(() => headerFooter.Text),
                UseFormat = Int(() => headerFooter.UseFormat),
                Visible = Int(() => headerFooter.Visible),
            };
        }

        private static PP.Comment BuildComment(Comment comment)
        {
            return new PP.Comment
            {
                Author = Str(() => comment.Author),
                AuthorInitials = Str(() => comment.AuthorInitials),
                Collapsed = Bool(() => comment.Collapsed),
                DateTime = DateStr(() => comment.DateTime),
                Left = SingleStr(() => comment.Left),
                ProviderID = Str(() => comment.ProviderID),
                Text = Str(() => comment.Text),
                Top = SingleStr(() => comment.Top),
                UserID = Str(() => comment.UserID),
            };
        }

        private static List<PP.Hyperlink> BuildHyperlinks(Hyperlinks hyperlinks)
        {
            var list = new List<PP.Hyperlink>();
            int count = hyperlinks.Count;

            for (int i = 1; i <= count; i++)
            {
                list.Add(BuildHyperlink(hyperlinks[i]));
            }

            return list;
        }

        private static PP.Hyperlink BuildHyperlink(Hyperlink hyperlink)
        {
            string? shapeName = string.Empty;
            string? text = string.Empty;

            try
            {
                // The parent chain is object-typed in the COM model; walk it late-bound
                // exactly like the VBA writer.
                MsoHyperlinkType type = hyperlink.Type;
                if (type == MsoHyperlinkType.msoHyperlinkShape)
                {
                    shapeName = Convert.ToString(((dynamic)hyperlink.UnderlyingObject).Parent.Parent.Name);
                }
                else if (type == MsoHyperlinkType.msoHyperlinkRange)
                {
                    text = Convert.ToString(((dynamic)hyperlink.UnderlyingObject).Parent.Parent.Text);
                    shapeName = Convert.ToString(((dynamic)hyperlink.UnderlyingObject).Parent.Parent.Parent.Parent.Name);
                }
            }
            catch
            {
                // Same as VBA: the derived ShapeName/Text stay empty when the parent chain fails.
            }

            return new PP.Hyperlink
            {
                Address = Str(() => hyperlink.Address),
                EmailSubject = Str(() => hyperlink.EmailSubject),
                ScreenTip = Str(() => hyperlink.ScreenTip),
                // ShowAndReturn is not exposed by the NetOffice Hyperlink wrapper; read it late-bound.
                ShowAndReturn = Int(() => ((dynamic)hyperlink.UnderlyingObject).ShowAndReturn),
                SubAddress = Str(() => hyperlink.SubAddress),
                TextToDisplay = Str(() => hyperlink.TextToDisplay),
                Type = Int(() => hyperlink.Type),
                ShapeName = shapeName ?? string.Empty,
                Text = text ?? string.Empty,
            };
        }

        private static List<PP.Tag> BuildTags(Tags tags)
        {
            var list = new List<PP.Tag>();
            int count = tags.Count;

            for (int i = 1; i <= count; i++)
            {
                int index = i;
                list.Add(new PP.Tag
                {
                    Name = Str(() => tags.Name(index)),
                    Value = Str(() => tags.Value(index)),
                });
            }

            return list;
        }

        // -- Shapes ---------------------------------------------------------------

        private static List<PP.Shape> BuildShapes(IEnumerable<Shape> shapes, string outputFilePath, MsOfficeJsonWriterOptions options)
        {
            var list = new List<PP.Shape>();

            foreach (Shape shape in shapes)
            {
                list.Add(BuildShape(shape, outputFilePath, options));
            }

            return list;
        }

        private static PP.Shape BuildShape(Shape shape, string outputFilePath, MsOfficeJsonWriterOptions options)
        {
            int? shapeType = Int(() => shape.Type);
            bool isPicture = shapeType == (int)MsoShapeType.msoPicture || shapeType == (int)MsoShapeType.msoLinkedPicture;

            var entity = new PP.Shape
            {
                AnimationSettings = Obj(() => BuildAnimationSettings(shape.AnimationSettings)),
                Fill = Obj(() => BuildFillFormat(shape.Fill)),
                Line = Obj(() => BuildLineFormat(shape.Line)),
                PictureFormat = isPicture ? Obj(() => BuildPictureFormat(shape.PictureFormat)) : null,
                PlaceholderFormat = Obj(() => BuildPlaceholderFormat(shape.PlaceholderFormat)),
                TextFrame = Obj<PP.TextFrame>(() =>
                {
                    string _ = shape.TextFrame.TextRange.Text;
                    return new PP.TextFrame { TextRange = BuildTextRange(shape.TextFrame.TextRange) };
                }),
                TextFrame2 = Obj<PP.TextFrame2>(() =>
                {
                    string _ = shape.TextFrame2.TextRange.Text;
                    return new PP.TextFrame2 { TextRange2 = BuildTextRange2(shape.TextFrame2.TextRange) };
                }),
                Image = isPicture ? BuildImage(shape, outputFilePath, options) : null,
                ActionSettings = Obj<List<PP.ActionSetting>>(() =>
                {
                    var list = new List<PP.ActionSetting>
                    {
                        BuildActionSetting(shape.ActionSettings[PpMouseActivation.ppMouseClick], (int)PpMouseActivation.ppMouseClick),
                        BuildActionSetting(shape.ActionSettings[PpMouseActivation.ppMouseOver], (int)PpMouseActivation.ppMouseOver),
                    };
                    return list;
                }),
                Tags = Obj(() => BuildTags(shape.Tags)),
                GroupItems = shapeType == (int)MsoShapeType.msoGroup
                    ? Obj<List<PP.Shape>>(() =>
                    {
                        var list = new List<PP.Shape>();
                        foreach (Shape item in shape.GroupItems) list.Add(BuildShape(item, outputFilePath, options));
                        return list;
                    })
                    : null,
                Id = Long(() => shape.Id),
                Name = Str(() => shape.Name),
                Type = shapeType,
                AutoShapeType = Int(() => shape.AutoShapeType),
                AlternativeText = Str(() => shape.AlternativeText),
                Title = Str(() => shape.Title),
                Left = SingleStr(() => shape.Left),
                Top = SingleStr(() => shape.Top),
                Width = SingleStr(() => shape.Width),
                Height = SingleStr(() => shape.Height),
                Rotation = SingleStr(() => shape.Rotation),
                ZOrderPosition = Int(() => shape.ZOrderPosition),
                Visible = Int(() => shape.Visible),
                BackgroundStyle = Int(() => shape.BackgroundStyle),
                BlackWhiteMode = Int(() => shape.BlackWhiteMode),
                Child = Int(() => shape.Child),
                ConnectionSiteCount = Long(() => shape.ConnectionSiteCount),
                Connector = Int(() => shape.Connector),
                Creator = Long(() => shape.Creator),
                Decorative = Int(() => shape.Decorative),
                // GraphicStyle, InkXML and the Has* companions below are newer Office
                // members the NetOffice wrappers do not expose; read them late-bound.
                GraphicStyle = Int(() => ((dynamic)shape.UnderlyingObject).GraphicStyle),
                HasChart = Int(() => shape.HasChart),
                HasInkXML = Int(() => ((dynamic)shape.UnderlyingObject).HasInkXML),
                HasSectionZoom = Int(() => ((dynamic)shape.UnderlyingObject).HasSectionZoom),
                HasSmartArt = Int(() => shape.HasSmartArt),
                HasTable = Int(() => shape.HasTable),
                HasTextFrame = Int(() => shape.HasTextFrame),
                HorizontalFlip = Int(() => shape.HorizontalFlip),
                InkXML = Str(() => ((dynamic)shape.UnderlyingObject).InkXML),
                IsNarration = Int(() => shape.IsNarration),
                LockAspectRatio = Int(() => shape.LockAspectRatio),
                ShapeStyle = Int(() => shape.ShapeStyle),
                VerticalFlip = Int(() => shape.VerticalFlip),
                Vertices = Str(() => shape.Vertices),
                MediaType = Int(() => shape.MediaType),
            };

            return entity;
        }

        private static ImageBlob? BuildImage(Shape shape, string outputFilePath, MsOfficeJsonWriterOptions options)
        {
            try
            {
                string tempPath = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".png");

                shape.Export(tempPath, PpShapeFormat.ppShapeFormatPNG);

                var blob = new ImageBlob { Extension = "png" };

                if (options.BlobOutput == BlobOutput_Enum.Base64)
                {
                    blob.Base64 = Convert.ToBase64String(File.ReadAllBytes(tempPath));
                }
                else
                {
                    string blobFolder = options.GetBlobFolderPath(outputFilePath);
                    Directory.CreateDirectory(blobFolder);

                    string fileName = shape.Id.ToString(CultureInfo.InvariantCulture) + ".png";
                    File.Copy(tempPath, Path.Combine(blobFolder, fileName), true);

                    blob.FileName = Path.GetFileName(blobFolder) + "/" + fileName;
                }

                File.Delete(tempPath);

                return blob;
            }
            catch
            {
                return null;
            }
        }

        // -- Shape formats ----------------------------------------------------------

        private static PP.AnimationSettings BuildAnimationSettings(AnimationSettings animationSettings)
        {
            return new PP.AnimationSettings
            {
                DimColor = Obj(() => BuildColorFormat(animationSettings.DimColor)),
                PlaySettings = Obj<PP.PlaySettings>(() =>
                {
                    // The PlaySettings getter throws for non-media shapes; like the VBA
                    // writer, the whole object is null in that case.
                    PlaySettings playSettings = animationSettings.PlaySettings;

                    return new PP.PlaySettings
                    {
                        ActionVerb = Str(() => playSettings.ActionVerb),
                        HideWhileNotPlaying = Int(() => playSettings.HideWhileNotPlaying),
                        LoopUntilStopped = Int(() => playSettings.LoopUntilStopped),
                        PauseAnimation = Int(() => playSettings.PauseAnimation),
                        PlayOnEntry = Int(() => playSettings.PlayOnEntry),
                        RewindMovie = Int(() => playSettings.RewindMovie),
                        StopAfterSlides = Long(() => playSettings.StopAfterSlides),
                    };
                }),
                SoundEffect = Obj(() => BuildSoundEffect(animationSettings.SoundEffect)),
                AdvanceMode = Int(() => animationSettings.AdvanceMode),
                AdvanceTime = SingleStr(() => animationSettings.AdvanceTime),
                AfterEffect = Int(() => animationSettings.AfterEffect),
                Animate = Int(() => animationSettings.Animate),
                AnimateBackground = Int(() => animationSettings.AnimateBackground),
                AnimateTextInReverse = Int(() => animationSettings.AnimateTextInReverse),
                AnimationOrder = Long(() => animationSettings.AnimationOrder),
                EntryEffect = Int(() => animationSettings.EntryEffect),
                TextLevelEffect = Int(() => animationSettings.TextLevelEffect),
                TextUnitEffect = Int(() => animationSettings.TextUnitEffect),
            };
        }

        private static PP.ActionSetting BuildActionSetting(ActionSetting actionSetting, int mouseActivation)
        {
            return new PP.ActionSetting
            {
                Hyperlink = Obj<PP.Hyperlink>(() =>
                {
                    if (actionSetting.Action != PpActionType.ppActionHyperlink) return null;
                    return BuildHyperlink(actionSetting.Hyperlink);
                }),
                SoundEffect = Obj(() => BuildSoundEffect(actionSetting.SoundEffect)),
                MouseActivation = mouseActivation,
                Action = Int(() => actionSetting.Action),
                ActionVerb = Str(() => actionSetting.ActionVerb),
                AnimateAction = Int(() => actionSetting.AnimateAction),
                Run = Str(() => actionSetting.Run),
                // ShowAndReturn is not exposed by the NetOffice ActionSetting wrapper; read it late-bound.
                ShowAndReturn = Int(() => ((dynamic)actionSetting.UnderlyingObject).ShowAndReturn),
                SlideShowName = Str(() => actionSetting.SlideShowName),
            };
        }

        private static PP.FillFormat BuildFillFormat(FillFormat fill)
        {
            return new PP.FillFormat
            {
                BackColor = Obj(() => BuildColorFormat(fill.BackColor)),
                ForeColor = Obj(() => BuildColorFormat(fill.ForeColor)),
                GradientAngle = SingleStr(() => fill.GradientAngle),
                GradientColorType = Int(() => fill.GradientColorType),
                GradientDegree = SingleStr(() => fill.GradientDegree),
                GradientStyle = Int(() => fill.GradientStyle),
                GradientVariant = Int(() => fill.GradientVariant),
                Pattern = Int(() => fill.Pattern),
                RotateWithObject = Int(() => fill.RotateWithObject),
                TextureAlignment = Int(() => fill.TextureAlignment),
                TextureHorizontalScale = SingleStr(() => fill.TextureHorizontalScale),
                TextureName = Str(() => fill.TextureName),
                TextureOffsetX = SingleStr(() => fill.TextureOffsetX),
                TextureOffsetY = SingleStr(() => fill.TextureOffsetY),
                TextureTile = Int(() => fill.TextureTile),
                TextureType = Int(() => fill.TextureType),
                TextureVerticalScale = SingleStr(() => fill.TextureVerticalScale),
                Transparency = SingleStr(() => fill.Transparency),
                Type = Int(() => fill.Type),
                Visible = Int(() => fill.Visible),
            };
        }

        private static PP.LineFormat BuildLineFormat(LineFormat line)
        {
            return new PP.LineFormat
            {
                BackColor = Obj(() => BuildColorFormat(line.BackColor)),
                ForeColor = Obj(() => BuildColorFormat(line.ForeColor)),
                BeginArrowheadLength = Int(() => line.BeginArrowheadLength),
                BeginArrowheadStyle = Int(() => line.BeginArrowheadStyle),
                BeginArrowheadWidth = Int(() => line.BeginArrowheadWidth),
                Creator = Long(() => line.Creator),
                DashStyle = Int(() => line.DashStyle),
                EndArrowheadLength = Int(() => line.EndArrowheadLength),
                EndArrowheadStyle = Int(() => line.EndArrowheadStyle),
                EndArrowheadWidth = Int(() => line.EndArrowheadWidth),
                InsetPen = Int(() => line.InsetPen),
                Pattern = Int(() => line.Pattern),
                Style = Int(() => line.Style),
                Transparency = SingleStr(() => line.Transparency),
                Visible = Int(() => line.Visible),
                Weight = SingleStr(() => line.Weight),
            };
        }

        private static PP.PictureFormat BuildPictureFormat(PictureFormat pictureFormat)
        {
            return new PP.PictureFormat
            {
                Brightness = SingleStr(() => pictureFormat.Brightness),
                ColorType = Int(() => pictureFormat.ColorType),
                Contrast = SingleStr(() => pictureFormat.Contrast),
                CropBottom = SingleStr(() => pictureFormat.CropBottom),
                CropLeft = SingleStr(() => pictureFormat.CropLeft),
                CropRight = SingleStr(() => pictureFormat.CropRight),
                CropTop = SingleStr(() => pictureFormat.CropTop),
                TransparencyColor = Long(() => pictureFormat.TransparencyColor),
                TransparentBackground = Int(() => pictureFormat.TransparentBackground),
            };
        }

        private static PP.PlaceholderFormat BuildPlaceholderFormat(PlaceholderFormat placeholderFormat)
        {
            return new PP.PlaceholderFormat
            {
                ContainedType = Int(() => placeholderFormat.ContainedType),
                Name = Str(() => placeholderFormat.Name),
                // Position is not exposed by the NetOffice PlaceholderFormat wrapper; read it late-bound.
                Position = Long(() => ((dynamic)placeholderFormat.UnderlyingObject).Position),
                Type = Int(() => placeholderFormat.Type),
            };
        }

        // -- Text -----------------------------------------------------------------

        private static PP.TextRange BuildTextRange(TextRange textRange)
        {
            return new PP.TextRange
            {
                Font = Obj(() => BuildFont(textRange.Font)),
                Length = Int(() => textRange.Length),
                Start = Int(() => textRange.Start),
                Text = Str(() => textRange.Text),
            };
        }

        private static PP.Font BuildFont(Font font)
        {
            return new PP.Font
            {
                Color = Obj(() => BuildColorFormat(font.Color)),
                AutoRotateNumbers = Int(() => font.AutoRotateNumbers),
                BaselineOffset = SingleStr(() => font.BaselineOffset),
                Bold = Int(() => font.Bold),
                Emboss = Int(() => font.Emboss),
                Italic = Int(() => font.Italic),
                Name = Str(() => font.Name),
                NameAscii = Str(() => font.NameAscii),
                NameComplexScript = Str(() => font.NameComplexScript),
                NameFarEast = Str(() => font.NameFarEast),
                NameOther = Str(() => font.NameOther),
                Shadow = Int(() => font.Shadow),
                Size = SingleStr(() => font.Size),
                Subscript = Int(() => font.Subscript),
                Superscript = Int(() => font.Superscript),
                Underline = Int(() => font.Underline),
            };
        }

        private static PP.TextRange2 BuildTextRange2(Office.TextRange2 textRange2)
        {
            return new PP.TextRange2
            {
                Font = Obj(() => BuildFont2(textRange2.Font)),
                Length = Int(() => textRange2.Length),
                Start = Int(() => textRange2.Start),
                Text = Str(() => textRange2.Text),
            };
        }

        private static PP.Font2 BuildFont2(Office.Font2 font2)
        {
            return new PP.Font2
            {
                Allcaps = Int(() => font2.Allcaps),
                // AutoRotateNumbers is not exposed by the NetOffice Font2 wrapper; read it late-bound.
                AutoRotateNumbers = Int(() => ((dynamic)font2.UnderlyingObject).AutoRotateNumbers),
                BaselineOffset = SingleStr(() => font2.BaselineOffset),
                Bold = Int(() => font2.Bold),
                Caps = Int(() => font2.Caps),
                DoubleStrikeThrough = Int(() => font2.DoubleStrikeThrough),
                Equalize = Int(() => font2.Equalize),
                Italic = Int(() => font2.Italic),
                Kerning = SingleStr(() => font2.Kerning),
                Name = Str(() => font2.Name),
                NameAscii = Str(() => font2.NameAscii),
                NameComplexScript = Str(() => font2.NameComplexScript),
                NameFarEast = Str(() => font2.NameFarEast),
                NameOther = Str(() => font2.NameOther),
                Size = SingleStr(() => font2.Size),
                Smallcaps = Int(() => font2.Smallcaps),
                SoftEdgeFormat = Int(() => font2.SoftEdgeFormat),
                Spacing = SingleStr(() => font2.Spacing),
                Strike = Int(() => font2.Strike),
                StrikeThrough = Int(() => font2.StrikeThrough),
                Subscript = Int(() => font2.Subscript),
                Superscript = Int(() => font2.Superscript),
                UnderlineStyle = Int(() => font2.UnderlineStyle),
                WordArtformat = Int(() => font2.WordArtformat),
            };
        }

        // -- Transitions, sounds, colors --------------------------------------------

        private static PP.SlideShowTransition BuildSlideShowTransition(SlideShowTransition transition)
        {
            return new PP.SlideShowTransition
            {
                SoundEffect = Obj(() => BuildSoundEffect(transition.SoundEffect)),
                AdvanceOnClick = Int(() => transition.AdvanceOnClick),
                AdvanceOnTime = Int(() => transition.AdvanceOnTime),
                AdvanceTime = SingleStr(() => transition.AdvanceTime),
                Duration = SingleStr(() => transition.Duration),
                EntryEffect = Int(() => transition.EntryEffect),
                Hidden = Int(() => transition.Hidden),
                LoopSoundUntilNext = Int(() => transition.LoopSoundUntilNext),
                Speed = Int(() => transition.Speed),
            };
        }

        private static PP.SoundEffect? BuildSoundEffect(SoundEffect soundEffect)
        {
            try
            {
                string _ = soundEffect.Name;
            }
            catch
            {
                return null;
            }

            return new PP.SoundEffect
            {
                Name = Str(() => soundEffect.Name),
                Type = Int(() => soundEffect.Type),
            };
        }

        private static PP.ColorFormat BuildColorFormat(ColorFormat colorFormat)
        {
            return new PP.ColorFormat
            {
                Brightness = SingleStr(() => colorFormat.Brightness),
                Creator = Long(() => colorFormat.Creator),
                ObjectThemeColor = Int(() => colorFormat.ObjectThemeColor),
                RGB = Int(() => colorFormat.RGB),
                SchemeColor = Int(() => colorFormat.SchemeColor),
                TintAndShade = SingleStr(() => colorFormat.TintAndShade),
                Type = Int(() => colorFormat.Type),
            };
        }

        private static PP.ColorScheme BuildColorScheme(ColorScheme colorScheme)
        {
            // PpColorSchemeIndex values (accessing an index that does not apply throws -> null).
            return new PP.ColorScheme
            {
                Accent1 = Str(() => colorScheme.Colors(PpColorSchemeIndex.ppAccent1).RGB),
                Accent2 = Str(() => colorScheme.Colors(PpColorSchemeIndex.ppAccent2).RGB),
                Accent3 = Str(() => colorScheme.Colors(PpColorSchemeIndex.ppAccent3).RGB),
                Background = Str(() => colorScheme.Colors(PpColorSchemeIndex.ppBackground).RGB),
                Fill = Str(() => colorScheme.Colors(PpColorSchemeIndex.ppFill).RGB),
                Foreground = Str(() => colorScheme.Colors(PpColorSchemeIndex.ppForeground).RGB),
                NotSchemeColor = Str(() => colorScheme.Colors(PpColorSchemeIndex.ppNotSchemeColor).RGB),
                SchemeColorMixed = Str(() => colorScheme.Colors(PpColorSchemeIndex.ppSchemeColorMixed).RGB),
                Shadow = Str(() => colorScheme.Colors(PpColorSchemeIndex.ppShadow).RGB),
                Title = Str(() => colorScheme.Colors(PpColorSchemeIndex.ppTitle).RGB),
            };
        }

        private static PP.ThemeColorScheme BuildThemeColorScheme(Office.ThemeColorScheme themeColorScheme)
        {
            // MsoThemeColorSchemeIndex values.
            return new PP.ThemeColorScheme
            {
                ThemeAccent1 = Str(() => themeColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeAccent1).RGB),
                ThemeAccent2 = Str(() => themeColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeAccent2).RGB),
                ThemeAccent3 = Str(() => themeColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeAccent3).RGB),
                ThemeAccent4 = Str(() => themeColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeAccent4).RGB),
                ThemeAccent5 = Str(() => themeColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeAccent5).RGB),
                ThemeAccent6 = Str(() => themeColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeAccent6).RGB),
                ThemeDark1 = Str(() => themeColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeDark1).RGB),
                ThemeDark2 = Str(() => themeColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeDark2).RGB),
                ThemeFollowedHyperlink = Str(() => themeColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeFollowedHyperlink).RGB),
                ThemeHyperlink = Str(() => themeColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeHyperlink).RGB),
                ThemeLight1 = Str(() => themeColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeLight1).RGB),
                ThemeLight2 = Str(() => themeColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeLight2).RGB),
            };
        }
    }
}
