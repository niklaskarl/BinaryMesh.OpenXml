using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;
using Charts = DocumentFormat.OpenXml.Drawing.Charts;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal sealed class OpenXmlGraphicFrameVisual : IOpenXmlVisual, IGraphicFrameVisual, IVisual
    {
        private readonly IOpenXmlVisualContainer container;

        private readonly GraphicFrame graphicFrame;

        public OpenXmlGraphicFrameVisual(IOpenXmlVisualContainer container, GraphicFrame graphicFrame)
        {
            this.container = container;
            this.graphicFrame = graphicFrame;
        }

        public uint Id => this.graphicFrame.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Id;

        public string Name => this.graphicFrame.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Name;

        public bool IsPlaceholder => this.graphicFrame.NonVisualGraphicFrameProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape?.HasCustomPrompt ?? false;

        public IShapeVisual AsShapeVisual()
        {
            return null;
        }

        public IGraphicFrameVisual AsGraphicFrameVisual()
        {
            return this;
        }

        public IGraphicFrameVisual SetOffset(long x, long y)
        {
            Transform transform = this.graphicFrame.Transform ?? (this.graphicFrame.Transform = new Transform());
            transform.Offset = new Drawing.Offset()
            {
                X = x,
                Y = y
            };

            return this;
        }

        public IGraphicFrameVisual SetExtents(long width, long height)
        {
            Transform transform = this.graphicFrame.Transform ?? (this.graphicFrame.Transform = new Transform());
            transform.Extents = new Drawing.Extents()
            {
                Cx = width,
                Cy = height
            };

            return this;
        }

        public IGraphicFrameVisual SetContent(IChartSpace chartSpace)
        {
            if (!(chartSpace is IOpenXmlChartSpace internalChartSpace))
            {
                throw new ArgumentException();
            }

            string id = this.container.Part.GetIdOfPart(internalChartSpace.ChartPart);
            if (id == null)
            {
                id = this.container.Part.CreateRelationshipToPartDefaultId(internalChartSpace.ChartPart);
            }

            Drawing.Graphic graphic = this.graphicFrame.Graphic ?? (this.graphicFrame.Graphic = new Drawing.Graphic());
            Charts.Chart chart = new Charts.Chart();
            chart.SetAttribute(new OpenXmlAttribute("r", "id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", id));

            graphic.GraphicData = new Drawing.GraphicData()
            {
                Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart"
            }
                .AppendChildFluent(chart);

            return this;
        }

        public OpenXmlElement CloneForSlide()
        {
            return this.graphicFrame.CloneNode(true);
        }

        IVisual IVisual.SetOffset(long x, long y)
        {
            return this.SetOffset(x, y);
        }

        IVisual IVisual.SetExtents(long width, long height)
        {
            return this.SetExtents(width, height);
        }
    }
}
