using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using JBC.ExploreTheWorld.DL.MsOfficeApi.MsOffice;
using PP = JBC.ExploreTheWorld.DL.MsOfficeApi.MsPowerPoint;
using static JBC.ExploreTheWorld.DL.MsOfficeApi.Dynamic_Impl.JsonWriters.ComTryGet;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.Dynamic_Impl.JsonWriters
{
    // Builds the canonical JBC.ExploreTheWorld.DL.MsOfficeApi.MsPowerPoint entity graph from a running
    // PowerPoint presentation COM object (late-bound or Interop PIA) and writes it with
    // MsOfficeJsonSerializer. The output matches MSO_MsPowerPoint_JsonWriter.bas.
    public static class MsPowerPointJsonWriter
    {
        // MsoShapeType / PpMouseActivation / PpActionType constants used late-bound.
        private const int msoGroup = 6;
        private const int msoLinkedPicture = 11;
        private const int msoPicture = 13;
        private const int msoHyperlinkRange = 0;
        private const int msoHyperlinkShape = 1;
        private const int ppMouseClick = 1;
        private const int ppMouseOver = 2;
        private const int ppActionHyperlink = 7;
        private const int ppShapeFormatPng = 2;

        // Accepts any PowerPoint Presentation COM object (late-bound or Interop PIA) —
        // used by the Dynamic repos and the Save-As-JSON helpers.
        public static void WritePresentationComToJsonFile(object presentationComObject, string outputFilePath,
            MsOfficeJsonWriterOptions? options = null, Action<string>? log = null)
        {
            options ??= new MsOfficeJsonWriterOptions();

            log?.Invoke($"Writing JSON: {outputFilePath}");

            dynamic com = presentationComObject;

            PP.Presentation entity = BuildPresentation(com, outputFilePath, options, log);

            MsOfficeJsonSerializer.WriteToFile(entity, outputFilePath);

            log?.Invoke($"Done: {outputFilePath}");
        }

        private static PP.Presentation BuildPresentation(dynamic p, string outputFilePath, MsOfficeJsonWriterOptions options, Action<string>? log)
        {
            var prs = new PP.Presentation();

            // Entities
            prs.SlideMaster = Obj(() => BuildMaster(p.SlideMaster, true, outputFilePath, options));
            prs.TitleMaster = Obj<PP.SlideMaster>(() =>
            {
                if (Convert.ToInt32(p.HasTitleMaster) != -1) return null;
                return BuildMaster(p.TitleMaster, false, outputFilePath, options);
            });
            prs.HandoutMaster = Obj(() => BuildMaster(p.HandoutMaster, false, outputFilePath, options));
            prs.NotesMaster = Obj(() => BuildMaster(p.NotesMaster, false, outputFilePath, options));
            prs.Coauthoring = Obj(() => new PP.Coauthoring
            {
                FavorServerEditsDuringMerge = Convert.ToBoolean(p.Coauthoring.FavorServerEditsDuringMerge),
                MergeMode = Convert.ToBoolean(p.Coauthoring.MergeMode),
                PendingUpdates = Convert.ToBoolean(p.Coauthoring.PendingUpdates),
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
                foreach (dynamic scheme in p.ColorSchemes) list.Add(BuildColorScheme(scheme));
                return list;
            });
            prs.CustomXmlParts = Obj<List<PP.CustomXmlPart>>(() =>
            {
                var list = new List<PP.CustomXmlPart>();
                foreach (dynamic part in p.CustomXMLParts)
                {
                    dynamic px = part;
                    list.Add(new PP.CustomXmlPart
                    {
                        BuiltIn = Bool(() => px.BuiltIn),
                        Id = Str(() => px.Id),
                        NamespaceURI = Str(() => px.NamespaceURI),
                        XML = Str(() => px.XML),
                    });
                }
                return list;
            });
            prs.Designs = Obj<List<PP.Design>>(() =>
            {
                var list = new List<PP.Design>();
                foreach (dynamic design in p.Designs)
                {
                    dynamic dx = design;
                    list.Add(new PP.Design
                    {
                        Name = Str(() => dx.Name),
                        Index = Long(() => dx.Index),
                        HasTitleMaster = Int(() => dx.HasTitleMaster),
                        Preserved = Int(() => dx.Preserved),
                    });
                }
                return list;
            });
            prs.ExtraColors = Obj<List<long>>(() =>
            {
                var list = new List<long>();
                int count = Convert.ToInt32(p.ExtraColors.Count);
                for (int i = 1; i <= count; i++) list.Add(Convert.ToInt64(p.ExtraColors[i]));
                return list;
            });
            prs.Fonts = Obj<List<PP.PresentationFont>>(() =>
            {
                var list = new List<PP.PresentationFont>();
                foreach (dynamic font in p.Fonts)
                {
                    dynamic fx = font;
                    list.Add(new PP.PresentationFont
                    {
                        Name = Str(() => fx.Name),
                        Embeddable = Int(() => fx.Embeddable),
                        Embedded = Int(() => fx.Embedded),
                    });
                }
                return list;
            });
            prs.Guides = Obj<List<PP.Guide>>(() =>
            {
                var list = new List<PP.Guide>();
                int count = Convert.ToInt32(p.Guides.Count);
                for (int i = 1; i <= count; i++)
                {
                    dynamic guide = p.Guides[i];
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
                int count = Convert.ToInt32(p.SectionProperties.Count);
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
                int count = Convert.ToInt32(p.Slides.Count);
                log?.Invoke($"  Writing {count} slides...");
                int index = 0;
                foreach (dynamic slide in p.Slides)
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
            // PowerPoint surfaces these MsoTriState properties as VARIANT_BOOL; the
            // canonical schema stores the MsoTriState value (msoTrue = -1), like VBA.
            prs.HasHandoutMaster = Int(() => Convert.ToBoolean(p.HasHandoutMaster) ? -1 : 0);
            prs.HasNotesMaster = Int(() => Convert.ToBoolean(p.HasNotesMaster) ? -1 : 0);
            prs.HasTitleMaster = Int(() => p.HasTitleMaster);
            prs.HasVBProject = Int(() => Convert.ToBoolean(p.HasVBProject) ? -1 : 0);
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

        private static PP.SlideMaster BuildMaster(dynamic master, bool isSlideMaster, string outputFilePath, MsOfficeJsonWriterOptions options)
        {
            var entity = new PP.SlideMaster
            {
                SlideShowTransition = Obj(() => BuildSlideShowTransition(master.SlideShowTransition)),
                SlideMasterLayouts = isSlideMaster
                    ? Obj<List<PP.SlideMasterLayout>>(() =>
                    {
                        var list = new List<PP.SlideMasterLayout>();
                        foreach (dynamic layout in master.CustomLayouts) list.Add(BuildSlideMasterLayout(layout, outputFilePath, options));
                        return list;
                    })
                    : null,
                Shapes = BuildShapes(master.Shapes, outputFilePath, options),
                Name = Str(() => master.Name),
            };

            return entity;
        }

        private static PP.SlideMasterLayout BuildSlideMasterLayout(dynamic layout, string outputFilePath, MsOfficeJsonWriterOptions options)
        {
            return new PP.SlideMasterLayout
            {
                SlideShowTransition = Obj(() => BuildSlideShowTransition(layout.SlideShowTransition)),
                Shapes = BuildShapes(layout.Shapes, outputFilePath, options),
                Hyperlinks = Obj<List<PP.Hyperlink>>(() =>
                {
                    int count = Convert.ToInt32(layout.Hyperlinks.Count);
                    if (count == 0) return null;
                    return BuildHyperlinks(layout.Hyperlinks);
                }),
                Name = Str(() => layout.Name),
                Index = Int(() => layout.Index),
            };
        }

        // -- Presentation sub-objects ---------------------------------------------

        private static PP.PageSetup BuildPageSetup(dynamic pageSetup)
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

        private static PP.PrintOptions BuildPrintOptions(dynamic printOptions)
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

        private static PP.SlideShowSettings BuildSlideShowSettings(dynamic settings)
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

        private static List<PP.BuiltInDocumentProperty> BuildDocumentProperties(dynamic properties)
        {
            var list = new List<PP.BuiltInDocumentProperty>();

            foreach (dynamic property in (IEnumerable)properties)
            {
                dynamic px = property;
                list.Add(new PP.BuiltInDocumentProperty
                {
                    Name = Str(() => px.Name),
                    Creator = Long(() => px.Creator),
                    LinkSource = Str(() => px.LinkSource),
                    LinkToContent = Int(() => px.LinkToContent),
                    Type = Int(() => px.Type),
                    Value = Str(() => px.Value),
                });
            }

            return list;
        }

        private static List<PP.CustomDocumentProperty> BuildCustomDocumentProperties(dynamic properties)
        {
            var list = new List<PP.CustomDocumentProperty>();

            foreach (dynamic property in (IEnumerable)properties)
            {
                dynamic px = property;
                list.Add(new PP.CustomDocumentProperty
                {
                    Name = Str(() => px.Name),
                    Creator = Long(() => px.Creator),
                    LinkSource = Str(() => px.LinkSource),
                    LinkToContent = Int(() => px.LinkToContent),
                    Type = Int(() => px.Type),
                    Value = Str(() => px.Value),
                });
            }

            return list;
        }

        private static List<PP.ContentTypeProperty> BuildContentTypeProperties(dynamic properties)
        {
            var list = new List<PP.ContentTypeProperty>();

            foreach (dynamic property in (IEnumerable)properties)
            {
                dynamic px = property;
                list.Add(new PP.ContentTypeProperty
                {
                    Id = Str(() => px.Id),
                    Name = Str(() => px.Name),
                    Creator = Long(() => px.Creator),
                    IsReadOnly = Bool(() => px.IsReadOnly),
                    IsRequired = Bool(() => px.IsRequired),
                    Type = Int(() => px.Type),
                    Value = Str(() => px.Value),
                });
            }

            return list;
        }

        // -- Slides ---------------------------------------------------------------

        private static PP.Slide BuildSlide(dynamic slide, string outputFilePath, MsOfficeJsonWriterOptions options)
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
                    int count = Convert.ToInt32(slide.Background.Count);
                    for (int i = 1; i <= count; i++) list.Add(BuildShape(slide.Background[i], outputFilePath, options));
                    return list;
                }),
                Shapes = BuildShapes(slide.Shapes, outputFilePath, options),
                Tags = Obj(() => BuildTags(slide.Tags)),
                Comments = Obj<List<PP.Comment>>(() =>
                {
                    int count = Convert.ToInt32(slide.Comments.Count);
                    if (count == 0) return null;
                    var list = new List<PP.Comment>();
                    for (int i = 1; i <= count; i++) list.Add(BuildComment(slide.Comments[i]));
                    return list;
                }),
                Hyperlinks = Obj<List<PP.Hyperlink>>(() =>
                {
                    int count = Convert.ToInt32(slide.Hyperlinks.Count);
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

        private static PP.HeadersFooters BuildHeadersFooters(dynamic headersFooters)
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

        private static PP.HeaderFooter BuildHeaderFooter(dynamic headerFooter)
        {
            return new PP.HeaderFooter
            {
                Format = Int(() => headerFooter.Format),
                Text = Str(() => headerFooter.Text),
                UseFormat = Int(() => headerFooter.UseFormat),
                Visible = Int(() => headerFooter.Visible),
            };
        }

        private static PP.Comment BuildComment(dynamic comment)
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

        private static List<PP.Hyperlink> BuildHyperlinks(dynamic hyperlinks)
        {
            var list = new List<PP.Hyperlink>();
            int count = Convert.ToInt32(hyperlinks.Count);

            for (int i = 1; i <= count; i++)
            {
                list.Add(BuildHyperlink(hyperlinks[i]));
            }

            return list;
        }

        private static PP.Hyperlink BuildHyperlink(dynamic hyperlink)
        {
            string? shapeName = string.Empty;
            string? text = string.Empty;

            try
            {
                int type = Convert.ToInt32(hyperlink.Type);
                if (type == msoHyperlinkShape)
                {
                    shapeName = Convert.ToString(hyperlink.Parent.Parent.Name);
                }
                else if (type == msoHyperlinkRange)
                {
                    text = Convert.ToString(hyperlink.Parent.Parent.Text);
                    shapeName = Convert.ToString(hyperlink.Parent.Parent.Parent.Parent.Name);
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
                ShowAndReturn = Int(() => hyperlink.ShowAndReturn),
                SubAddress = Str(() => hyperlink.SubAddress),
                TextToDisplay = Str(() => hyperlink.TextToDisplay),
                Type = Int(() => hyperlink.Type),
                ShapeName = shapeName ?? string.Empty,
                Text = text ?? string.Empty,
            };
        }

        private static List<PP.Tag> BuildTags(dynamic tags)
        {
            var list = new List<PP.Tag>();
            int count = Convert.ToInt32(tags.Count);

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

        private static List<PP.Shape> BuildShapes(dynamic shapes, string outputFilePath, MsOfficeJsonWriterOptions options)
        {
            var list = new List<PP.Shape>();

            foreach (dynamic shape in shapes)
            {
                list.Add(BuildShape(shape, outputFilePath, options));
            }

            return list;
        }

        private static PP.Shape BuildShape(dynamic shape, string outputFilePath, MsOfficeJsonWriterOptions options)
        {
            int? shapeType = Int(() => shape.Type);
            bool isPicture = shapeType == msoPicture || shapeType == msoLinkedPicture;

            var entity = new PP.Shape
            {
                AnimationSettings = Obj(() => BuildAnimationSettings(shape.AnimationSettings)),
                Fill = Obj(() => BuildFillFormat(shape.Fill)),
                Line = Obj(() => BuildLineFormat(shape.Line)),
                PictureFormat = isPicture ? Obj(() => BuildPictureFormat(shape.PictureFormat)) : null,
                PlaceholderFormat = Obj(() => BuildPlaceholderFormat(shape.PlaceholderFormat)),
                TextFrame = Obj<PP.TextFrame>(() =>
                {
                    string _ = Convert.ToString(shape.TextFrame.TextRange.Text);
                    return new PP.TextFrame { TextRange = BuildTextRange(shape.TextFrame.TextRange) };
                }),
                TextFrame2 = Obj<PP.TextFrame2>(() =>
                {
                    string _ = Convert.ToString(shape.TextFrame2.TextRange.Text);
                    return new PP.TextFrame2 { TextRange2 = BuildTextRange2(shape.TextFrame2.TextRange) };
                }),
                Image = isPicture ? BuildImage(shape, outputFilePath, options) : null,
                ActionSettings = Obj<List<PP.ActionSetting>>(() =>
                {
                    var list = new List<PP.ActionSetting>
                    {
                        BuildActionSetting(shape.ActionSettings[ppMouseClick], ppMouseClick),
                        BuildActionSetting(shape.ActionSettings[ppMouseOver], ppMouseOver),
                    };
                    return list;
                }),
                Tags = Obj(() => BuildTags(shape.Tags)),
                GroupItems = shapeType == msoGroup
                    ? Obj<List<PP.Shape>>(() =>
                    {
                        var list = new List<PP.Shape>();
                        foreach (dynamic item in shape.GroupItems) list.Add(BuildShape(item, outputFilePath, options));
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
                GraphicStyle = Int(() => shape.GraphicStyle),
                HasChart = Int(() => shape.HasChart),
                HasInkXML = Int(() => shape.HasInkXML),
                HasSectionZoom = Int(() => shape.HasSectionZoom),
                HasSmartArt = Int(() => shape.HasSmartArt),
                HasTable = Int(() => shape.HasTable),
                HasTextFrame = Int(() => shape.HasTextFrame),
                HorizontalFlip = Int(() => shape.HorizontalFlip),
                InkXML = Str(() => shape.InkXML),
                IsNarration = Int(() => shape.IsNarration),
                LockAspectRatio = Int(() => shape.LockAspectRatio),
                ShapeStyle = Int(() => shape.ShapeStyle),
                VerticalFlip = Int(() => shape.VerticalFlip),
                Vertices = Str(() => shape.Vertices),
                MediaType = Int(() => shape.MediaType),
            };

            return entity;
        }

        private static ImageBlob? BuildImage(dynamic shape, string outputFilePath, MsOfficeJsonWriterOptions options)
        {
            try
            {
                string tempPath = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".png");

                shape.Export(tempPath, ppShapeFormatPng);

                var blob = new ImageBlob { Extension = "png" };

                if (options.BlobOutput == BlobOutput_Enum.Base64)
                {
                    blob.Base64 = Convert.ToBase64String(File.ReadAllBytes(tempPath));
                }
                else
                {
                    string blobFolder = options.GetBlobFolderPath(outputFilePath);
                    Directory.CreateDirectory(blobFolder);

                    string fileName = Convert.ToString(shape.Id) + ".png";
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

        private static PP.AnimationSettings BuildAnimationSettings(dynamic animationSettings)
        {
            return new PP.AnimationSettings
            {
                DimColor = Obj(() => BuildColorFormat(animationSettings.DimColor)),
                PlaySettings = Obj<PP.PlaySettings>(() =>
                {
                    // The PlaySettings getter throws for non-media shapes; like the VBA
                    // writer, the whole object is null in that case.
                    dynamic playSettings = animationSettings.PlaySettings;

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

        private static PP.ActionSetting BuildActionSetting(dynamic actionSetting, int mouseActivation)
        {
            return new PP.ActionSetting
            {
                Hyperlink = Obj<PP.Hyperlink>(() =>
                {
                    if (Convert.ToInt32(actionSetting.Action) != ppActionHyperlink) return null;
                    return BuildHyperlink(actionSetting.Hyperlink);
                }),
                SoundEffect = Obj(() => BuildSoundEffect(actionSetting.SoundEffect)),
                MouseActivation = mouseActivation,
                Action = Int(() => actionSetting.Action),
                ActionVerb = Str(() => actionSetting.ActionVerb),
                AnimateAction = Int(() => actionSetting.AnimateAction),
                Run = Str(() => actionSetting.Run),
                ShowAndReturn = Int(() => actionSetting.ShowAndReturn),
                SlideShowName = Str(() => actionSetting.SlideShowName),
            };
        }

        private static PP.FillFormat BuildFillFormat(dynamic fill)
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

        private static PP.LineFormat BuildLineFormat(dynamic line)
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

        private static PP.PictureFormat BuildPictureFormat(dynamic pictureFormat)
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

        private static PP.PlaceholderFormat BuildPlaceholderFormat(dynamic placeholderFormat)
        {
            return new PP.PlaceholderFormat
            {
                ContainedType = Int(() => placeholderFormat.ContainedType),
                Name = Str(() => placeholderFormat.Name),
                Position = Long(() => placeholderFormat.Position),
                Type = Int(() => placeholderFormat.Type),
            };
        }

        // -- Text -----------------------------------------------------------------

        private static PP.TextRange BuildTextRange(dynamic textRange)
        {
            return new PP.TextRange
            {
                Font = Obj(() => BuildFont(textRange.Font)),
                Length = Int(() => textRange.Length),
                Start = Int(() => textRange.Start),
                Text = Str(() => textRange.Text),
            };
        }

        private static PP.Font BuildFont(dynamic font)
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

        private static PP.TextRange2 BuildTextRange2(dynamic textRange2)
        {
            return new PP.TextRange2
            {
                Font = Obj(() => BuildFont2(textRange2.Font)),
                Length = Int(() => textRange2.Length),
                Start = Int(() => textRange2.Start),
                Text = Str(() => textRange2.Text),
            };
        }

        private static PP.Font2 BuildFont2(dynamic font2)
        {
            return new PP.Font2
            {
                Allcaps = Int(() => font2.Allcaps),
                AutoRotateNumbers = Int(() => font2.AutoRotateNumbers),
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
                WordArtformat = Int(() => font2.WordArtFormat),
            };
        }

        // -- Transitions, sounds, colors --------------------------------------------

        private static PP.SlideShowTransition BuildSlideShowTransition(dynamic transition)
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

        private static PP.SoundEffect? BuildSoundEffect(dynamic soundEffect)
        {
            try
            {
                string _ = Convert.ToString(soundEffect.Name);
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

        private static PP.ColorFormat BuildColorFormat(dynamic colorFormat)
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

        private static PP.ColorScheme BuildColorScheme(dynamic colorScheme)
        {
            // PpColorSchemeIndex values (accessing an index that does not apply throws -> null).
            return new PP.ColorScheme
            {
                Accent1 = Str(() => colorScheme.Colors(6).RGB),
                Accent2 = Str(() => colorScheme.Colors(7).RGB),
                Accent3 = Str(() => colorScheme.Colors(8).RGB),
                Background = Str(() => colorScheme.Colors(1).RGB),
                Fill = Str(() => colorScheme.Colors(5).RGB),
                Foreground = Str(() => colorScheme.Colors(2).RGB),
                NotSchemeColor = Str(() => colorScheme.Colors(0).RGB),
                SchemeColorMixed = Str(() => colorScheme.Colors(-2).RGB),
                Shadow = Str(() => colorScheme.Colors(3).RGB),
                Title = Str(() => colorScheme.Colors(4).RGB),
            };
        }

        private static PP.ThemeColorScheme BuildThemeColorScheme(dynamic themeColorScheme)
        {
            // MsoThemeColorSchemeIndex values.
            return new PP.ThemeColorScheme
            {
                ThemeAccent1 = Str(() => themeColorScheme.Colors(5).RGB),
                ThemeAccent2 = Str(() => themeColorScheme.Colors(6).RGB),
                ThemeAccent3 = Str(() => themeColorScheme.Colors(7).RGB),
                ThemeAccent4 = Str(() => themeColorScheme.Colors(8).RGB),
                ThemeAccent5 = Str(() => themeColorScheme.Colors(9).RGB),
                ThemeAccent6 = Str(() => themeColorScheme.Colors(10).RGB),
                ThemeDark1 = Str(() => themeColorScheme.Colors(1).RGB),
                ThemeDark2 = Str(() => themeColorScheme.Colors(3).RGB),
                ThemeFollowedHyperlink = Str(() => themeColorScheme.Colors(12).RGB),
                ThemeHyperlink = Str(() => themeColorScheme.Colors(11).RGB),
                ThemeLight1 = Str(() => themeColorScheme.Colors(2).RGB),
                ThemeLight2 = Str(() => themeColorScheme.Colors(4).RGB),
            };
        }
    }
}
