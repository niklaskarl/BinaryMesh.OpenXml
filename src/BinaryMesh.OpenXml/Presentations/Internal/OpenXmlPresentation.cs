using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Charts = DocumentFormat.OpenXml.Drawing.Charts;

using BinaryMesh.OpenXml.Helpers;
using BinaryMesh.OpenXml.Spreadsheets;
using System.Threading.Tasks;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal class OpenXmlPresentation : IOpenXmlPresentation, IPresentation, IDisposable
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

        public IReadOnlyList<ISlideMaster> SlideMasters => new EnumerableList<SlideMasterId, ISlideMaster>(
            this.presentationPart.Presentation.SlideMasterIdList.Elements<SlideMasterId>(),
            sm => new OpenXmlSlideMaster(this, this.presentationPart.GetPartById(sm.RelationshipId) as SlideMasterPart)
        );

        public IReadOnlyList<ISlide> Slides => new EnumerableList<SlideId, ISlide>(
            this.presentationPart.Presentation.SlideIdList.Elements<SlideId>(),
            slideId => new OpenXmlSlide(this, this.presentationPart.GetPartById(slideId.RelationshipId) as SlidePart)
        );

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
    }
}
