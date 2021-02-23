using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

using Drawing = DocumentFormat.OpenXml.Drawing;

namespace BinaryMesh.OpenXml.Presentations
{
    internal static class PresentationBuilder
    {
        public static PresentationDocument BuildDefault(PresentationDocument presentationDocument)
        {
            PresentationBuilder.BuildPresentationScaffold(presentationDocument);

            PresentationPart presentationPart = presentationDocument.PresentationPart;

            ThemeApplier.ApplyTheme(presentationPart, DefaultThemeBuilder.BuildDefaultTheme());
            ThemeApplier.ApplyTableStyles(presentationPart, DefaultThemeBuilder.BuildDefaultTableStyleList());
            ThemeApplier.AppendSlideMaster(presentationPart, DefaultSlideMasterBuilder.BuildDefaultSlideMaster(), DefaultSlideMasterBuilder.BuildDefaultSlideLayouts());

            return presentationDocument;
        }

        public static PresentationDocument BuildFromThemeDocument(PresentationDocument presentationDocument, PresentationDocument themePresentationDocument)
        {
            PresentationBuilder.BuildPresentationScaffold(presentationDocument);

            PresentationPart presentationPart = presentationDocument.PresentationPart;
            PresentationPart themePresentationPart = themePresentationDocument.PresentationPart;

            ThemeApplier.ApplyTheme(presentationPart, themePresentationPart.ThemePart.Theme);
            ThemeApplier.ApplyTableStyles(presentationPart, themePresentationPart.TableStylesPart.TableStyleList);

            foreach (SlideMasterPart themeSlideMasterPart in themePresentationPart.SlideMasterParts)
            {
                // ThemeApplier.AppendSlideMaster(presentationPart, themeSlideMasterPart.SlideMaster, ThemeApplier.ExtractSlideLayouts(themeSlideMasterPart));
                ThemeApplier.AppendSlideMasterFromPart(presentationPart, themePresentationPart, themeSlideMasterPart);
            }

            return presentationDocument;
        }

        private static PresentationDocument BuildPresentationScaffold(PresentationDocument presentationDocument)
        {
            presentationDocument.DeletePartsRecursivelyOfType<PresentationPart>();
            presentationDocument.DeletePartsRecursivelyOfType<CoreFilePropertiesPart>();
            presentationDocument.DeletePartsRecursivelyOfType<ExtendedFilePropertiesPart>();

            presentationDocument.ChangeIdOfPart(presentationDocument.AddPresentationPart(), "rId1");
            PresentationBuilder.MakeValidPresentationPart(presentationDocument.PresentationPart);

            presentationDocument.ChangeIdOfPart(presentationDocument.AddCoreFilePropertiesPart(), "rId2");
            PresentationBuilder.MakeValidCoreFilePropertiesPart(presentationDocument.CoreFilePropertiesPart);

            presentationDocument.ChangeIdOfPart(presentationDocument.AddExtendedFilePropertiesPart(), "rId3");
            PresentationBuilder.MakeValidExtendedFilePropertiesPart(presentationDocument.ExtendedFilePropertiesPart);

            return presentationDocument;
        }

        private static CoreFilePropertiesPart MakeValidCoreFilePropertiesPart(CoreFilePropertiesPart coreFilePropertiesPart)
        {
            byte[] blob = System.Text.Encoding.UTF8.GetBytes("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"><cp:lastModifiedBy>Niklas Karl</cp:lastModifiedBy><cp:revision>1</cp:revision><dcterms:modified xsi:type=\"dcterms:W3CDTF\">2021-02-18T15:56:56Z</dcterms:modified></cp:coreProperties>");

            using (MemoryStream stream = new MemoryStream(blob))
            {
                coreFilePropertiesPart.FeedData(stream);
            }

            return coreFilePropertiesPart;
        }

        private static ExtendedFilePropertiesPart MakeValidExtendedFilePropertiesPart(ExtendedFilePropertiesPart extendedFilePropertiesPart)
        {
            byte[] blob = System.Text.Encoding.UTF8.GetBytes("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\"></Properties>");

            using (MemoryStream stream = new MemoryStream(blob))
            {
                extendedFilePropertiesPart.FeedData(stream);
            }

            return extendedFilePropertiesPart;
        }

        private static PresentationPart MakeValidPresentationPart(PresentationPart presentationPart)
        {
            //
            // Presentation
            //
            PresentationBuilder.MakeValidPresentation(
                presentationPart.Presentation ?? (presentationPart.Presentation = new Presentation())
            ).Save(presentationPart);

            //
            // ViewPropertiesPart
            //
            PresentationBuilder.MakeValidViewPropertiesPart(
                presentationPart.ViewPropertiesPart ?? presentationPart.AddNewPartDefaultId<ViewPropertiesPart>()
            );

            //
            // PresentationPropertiesPart
            //
            PresentationBuilder.MakeValidPresentationPropertiesPart(
                presentationPart.PresentationPropertiesPart ?? presentationPart.AddNewPartDefaultId<PresentationPropertiesPart>()
            );

            return presentationPart;
        }

        private static Presentation MakeValidPresentation(Presentation presentation)
        {
            if (presentation.SlideMasterIdList == null)
            {
                presentation.SlideMasterIdList = new SlideMasterIdList();
            }

            if (presentation.SlideSize == null)
            {
                presentation.SlideSize = new SlideSize()
                {
                    Cx = 12192000,
                    Cy = 6858000
                };
            }

            if (presentation.NotesSize == null)
            {
                presentation.NotesSize = new NotesSize()
                {
                    Cx = 6858000,
                    Cy = 9144000
                };
            }

            if (presentation.DefaultTextStyle == null)
            {
                presentation.DefaultTextStyle = new DefaultTextStyle()
                {
                    DefaultParagraphProperties = new Drawing.DefaultParagraphProperties()
                };
            }

            return presentation;
        }

        private static ViewPropertiesPart MakeValidViewPropertiesPart(ViewPropertiesPart viewPropertiesPart)
        {
            PresentationBuilder.MakeValidViewProperties(
                viewPropertiesPart.ViewProperties ?? (viewPropertiesPart.ViewProperties = new ViewProperties())
            ).Save(viewPropertiesPart);

            return viewPropertiesPart;
        }

        private static ViewProperties MakeValidViewProperties(ViewProperties viewProperties)
        {
            viewProperties.NormalViewProperties = new NormalViewProperties()
            {
                HorizontalBarState = SplitterBarStateValues.Maximized,
                RestoredLeft = new RestoredLeft() { Size = 15987, AutoAdjust = false },
                RestoredTop = new RestoredTop() { Size = 94660 },
            };

            viewProperties.SlideViewProperties = new SlideViewProperties()
            {
                CommonSlideViewProperties = new CommonSlideViewProperties()
                {
                    SnapToGrid = false,
                    CommonViewProperties = new CommonViewProperties()
                    {
                        VariableScale = true,
                        ScaleFactor = new ScaleFactor()
                        {
                            ScaleX = new Drawing.ScaleX() { Numerator = 114, Denominator = 100 },
                            ScaleY = new Drawing.ScaleY() { Numerator = 114, Denominator = 100 }
                        },
                        Origin = new Origin()
                        {
                            X = 414,
                            Y = 102
                        }
                    },
                    GuideList = new GuideList()
                }
            };

            viewProperties.NotesTextViewProperties = new NotesTextViewProperties()
            {
                CommonViewProperties = new CommonViewProperties()
                {
                    ScaleFactor = new ScaleFactor()
                    {
                        ScaleX = new Drawing.ScaleX() { Numerator = 1, Denominator = 1 },
                        ScaleY = new Drawing.ScaleY() { Numerator = 1, Denominator = 1 }
                    },
                    Origin = new Origin()
                    {
                        X = 0,
                        Y = 0
                    }
                }
            };

            viewProperties.GridSpacing = new GridSpacing()
            {
                Cx = 72008,
                Cy = 72008
            };

            return viewProperties;
        }

        private static PresentationPropertiesPart MakeValidPresentationPropertiesPart(PresentationPropertiesPart presentationPropertiesPart)
        {
            PresentationBuilder.MakeValidPresentationProperties(
                presentationPropertiesPart.PresentationProperties ?? (presentationPropertiesPart.PresentationProperties = new PresentationProperties())
            ).Save(presentationPropertiesPart);

            return presentationPropertiesPart;
        }

        private static PresentationProperties MakeValidPresentationProperties(PresentationProperties presentationProperties)
        {
            // TODO optional
            return presentationProperties;
        }
    }
}
