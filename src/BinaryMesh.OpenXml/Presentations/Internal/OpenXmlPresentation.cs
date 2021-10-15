using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;

using BinaryMesh.OpenXml.Internal;
using BinaryMesh.OpenXml.Helpers;
using BinaryMesh.OpenXml.Tables;
using BinaryMesh.OpenXml.Tables.Internal;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal class OpenXmlPresentation : IOpenXmlPresentation, IPresentation, IOpenXmlDocument, IDisposable
    {
        private readonly Stream stream;

        private readonly PresentationDocument presentationDocument;

        private readonly PresentationPart presentationPart;

        public OpenXmlPresentation(Stream template)
        {
            /*this.stream = new MemoryStream();
            this.presentationDocument = PresentationDocument.Create(stream, PresentationDocumentType.Presentation);
            if (template != null)
            {
                using (PresentationDocument themePresentationDocument = PresentationDocument.Open(template, false))
                {
                    PresentationBuilder.BuildFromThemeDocument(this.presentationDocument, themePresentationDocument);
                }
            }
            else
            {
                PresentationBuilder.BuildDefault(this.presentationDocument);
            }

            this.presentationPart = this.presentationDocument.PresentationPart;*/

            this.stream = new MemoryStream();

            if (template != null)
            {
                template.CopyTo(this.stream);
                this.stream.Seek(0, SeekOrigin.Begin);
            }

            this.presentationDocument = PresentationDocument.Open(this.stream, true);
            this.presentationPart = this.presentationDocument.PresentationPart;
        }

        public PresentationPart PresentationPart => this.presentationPart;

        public IOpenXmlTextStyle DefaultTextStyle => new PresentationDefaultTextStyle(this);

        public IOpenXmlTheme Theme => new OpenXmlTheme(this, this.presentationPart.ThemePart);

        public IReadOnlyList<ISlideMaster> SlideMasters => new EnumerableList<SlideMasterId, ISlideMaster>(
            this.presentationPart.Presentation.SlideMasterIdList.Elements<SlideMasterId>(),
            sm => new OpenXmlSlideMaster(this, this.presentationPart.GetPartById(sm.RelationshipId) as SlideMasterPart)
        );

        public IReadOnlyList<ISlide> Slides => new EnumerableList<SlideId, ISlide>(
            this.presentationPart.Presentation.SlideIdList.Elements<SlideId>(),
            slideId => new OpenXmlSlide(this, this.presentationPart.GetPartById(slideId.RelationshipId) as SlidePart)
        );

        public ITableStyleCollection TableStyles => new OpenXmlTableStyleCollection(this.GetTableStylesPart);

        public ISlide InsertSlide(ISlideLayout slideLayout)
        {
            return this.InsertSlide(slideLayout, 0);
        }

        public ISlide InsertSlide(ISlideLayout slideLayout, int index)
        {
            if (!(slideLayout is IOpenXmlSlideLayout internalSlideLayout))
            {
                throw new ArgumentException();
            }

            Slide slide = new Slide()
            {
                // CommonSlideData = slideLayoutRef.SlideLayoutPart.SlideLayout.CommonSlideData.CloneNode(true) as CommonSlideData
                CommonSlideData = new CommonSlideData()
                {
                    ShapeTree = new ShapeTree()
                    {
                        NonVisualGroupShapeProperties = internalSlideLayout.SlideLayoutPart.SlideLayout.CommonSlideData.ShapeTree.NonVisualGroupShapeProperties.CloneNode(true) as NonVisualGroupShapeProperties,
                        GroupShapeProperties = internalSlideLayout.SlideLayoutPart.SlideLayout.CommonSlideData.ShapeTree.GroupShapeProperties.CloneNode(true) as GroupShapeProperties
                    }
                }
            };

            slide.CommonSlideData.ShapeTree.Append(
                internalSlideLayout.SlideLayoutPart.SlideLayout.CommonSlideData.ShapeTree
                    .Select(element => OpenXmlVisualFactory.TryCreateVisual(internalSlideLayout, element, out IOpenXmlVisual visual) ? visual : null).Where(visual => visual != null)
                    .Where(visual => visual?.IsPlaceholder ?? false)
                    .Select(visual => visual.CloneForSlide())
            );

            SlidePart slidePart = this.presentationPart.AddNewPart<SlidePart>();
            slide.Save(slidePart);

            slidePart.CreateRelationshipToPartDefaultId(internalSlideLayout.SlideLayoutPart);

            if (this.presentationPart.Presentation.SlideIdList == null)
            {
                this.presentationPart.Presentation.SlideIdList = new SlideIdList();
            }

            SlideIdList slideIdList = this.presentationPart.Presentation.SlideIdList;
            SlideId refSlideId = slideIdList.Elements<SlideId>().Take(index).LastOrDefault();

            uint id = refSlideId?.Id ?? 256;
            SlideId slideId;
            if (refSlideId != null)
            {
                slideId = slideIdList.InsertAfter(
                    new SlideId() { Id = ++id, RelationshipId = presentationPart.GetIdOfPart(slidePart) },
                    refSlideId
                );
            }
            else
            {
                slideId = slideIdList.PrependChild(
                    new SlideId() { Id = ++id, RelationshipId = presentationPart.GetIdOfPart(slidePart) }
                );
            }
            
            for (slideId = slideId.NextSibling<SlideId>(); slideId != null; slideId = slideId.NextSibling<SlideId>())
            {
                slideId.Id = ++id;
            }

            return new OpenXmlSlide(this, slidePart);
        }

        public void Close(Stream destination)
        {
            this.presentationDocument.Close();
            this.stream.Position = 0;
            this.stream.CopyTo(destination);
            this.Dispose();
        }

        public async Task CloseAsync(Stream destination)
        {
            this.presentationDocument.Close();
            this.stream.Position = 0;
            await this.stream.CopyToAsync(destination);
            this.Dispose();
        }

        public void Dispose()
        {
            this.presentationDocument.Dispose();
            this.stream.Dispose();
        }

        private TableStylesPart GetTableStylesPart(bool create)
        {
            TableStylesPart part = this.presentationPart.GetPartsOfType<TableStylesPart>().FirstOrDefault();
            if (part == null && create)
            {
                part = this.presentationPart.AddNewPartDefaultId<TableStylesPart>();
                part.TableStyleList = new Drawing.TableStyleList();
            }

            return part;
        }

        private class PresentationDefaultTextStyle : IOpenXmlTextStyle
        {
            private OpenXmlPresentation presentation;

            public PresentationDefaultTextStyle(OpenXmlPresentation presentation)
            {
                this.presentation = presentation;
            }

            public IOpenXmlParagraphTextStyle GetParagraphTextStyle(int level)
            {
                DefaultTextStyle style = this.presentation.presentationPart.Presentation.DefaultTextStyle;
                Drawing.TextParagraphPropertiesType properties = null;
                switch (level)
                {
                    case 0:
                        properties = style.DefaultParagraphProperties;
                        break;
                    case 1:
                        properties = style.Level1ParagraphProperties;
                        break;
                    case 2:
                        properties = style.Level2ParagraphProperties;
                        break;
                    case 3:
                        properties = style.Level3ParagraphProperties;
                        break;
                    case 4:
                        properties = style.Level4ParagraphProperties;
                        break;
                    case 5:
                        properties = style.Level5ParagraphProperties;
                        break;
                    case 6:
                        properties = style.Level6ParagraphProperties;
                        break;
                    case 7:
                        properties = style.Level7ParagraphProperties;
                        break;
                    case 8:
                        properties = style.Level8ParagraphProperties;
                        break;
                    case 9:
                        properties = style.Level9ParagraphProperties;
                        break;
                }

                return new OpenXmlParagraphTextStyle(properties);
            }
        }
    }
}
