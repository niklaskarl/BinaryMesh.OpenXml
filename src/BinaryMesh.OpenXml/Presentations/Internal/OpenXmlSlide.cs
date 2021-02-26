using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Charts = DocumentFormat.OpenXml.Drawing.Charts;

using BinaryMesh.OpenXml.Helpers;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal sealed class OpenXmlSlide : IOpenXmlSlide, IOpenXmlVisualContainer, ISlide
    {
        private readonly IOpenXmlPresentation presentation;

        private readonly SlidePart slidePart;

        public OpenXmlSlide(IOpenXmlPresentation presentation, SlidePart slidePart)
        {
            this.presentation = presentation;
            this.slidePart = slidePart;
        }

        public OpenXmlPart Part => this.slidePart;

        public SlidePart SlidePart => this.slidePart;

        public int Index => throw new NotImplementedException();

        public KeyedReadOnlyList<string, IVisual> VisualTree => new EnumerableKeyedList<IOpenXmlVisual, string, IVisual>(
            this.slidePart.Slide.CommonSlideData.ShapeTree.Select(element => OpenXmlVisualFactory.TryCreateVisual(this, element, out IOpenXmlVisual visual) ? visual : null).Where(visual => visual != null),
            visual => visual.Name,
            visual => visual
        );

        public IShapeVisual AppendShapeVisual(string name)
        {
            throw new NotImplementedException();
        }

        public IGraphicFrameVisual AppendGraphicFrameVisual(string name)
        {
            ShapeTree shapeTree = this.slidePart.Slide.CommonSlideData.ShapeTree;
            GraphicFrame graphicFrame = shapeTree.AppendChild(new GraphicFrame()
            {
                NonVisualGraphicFrameProperties = new NonVisualGraphicFrameProperties()
                {
                    NonVisualDrawingProperties = new NonVisualDrawingProperties()
                    {
                        Name = name,
                        Id = 6 // TODO: calculate
                    },
                    NonVisualGraphicFrameDrawingProperties = new NonVisualGraphicFrameDrawingProperties(),
                    ApplicationNonVisualDrawingProperties = new ApplicationNonVisualDrawingProperties()
                }
            });

            return new OpenXmlGraphicFrameVisual(this, graphicFrame);
        }

        public IChartSpace CreateChartSpace()
        {
            ChartPart chartPart = this.slidePart.AddNewPartDefaultId<ChartPart>();
            chartPart.ChartSpace = new Charts.ChartSpace()
            {
                Date1904 = new Charts.Date1904() { Val = false }
            }
                .AppendChildFluent(new Charts.Chart());

            return new OpenXmlChartSpace(chartPart);
        }
    }
}
