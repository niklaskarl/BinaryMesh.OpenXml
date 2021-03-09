using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;
using Charts = DocumentFormat.OpenXml.Drawing.Charts;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal class OpenXmlGraphicFrameVisual : IOpenXmlVisual, IVisual
    {
        protected readonly IOpenXmlVisualContainer container;

        protected readonly GraphicFrame graphicFrame;

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

        public IVisual SetOffset(long x, long y)
        {
            Transform transform = this.graphicFrame.Transform ?? (this.graphicFrame.Transform = new Transform());
            transform.Offset = new Drawing.Offset()
            {
                X = x,
                Y = y
            };

            return this;
        }

        public IVisual SetExtents(long width, long height)
        {
            Transform transform = this.graphicFrame.Transform ?? (this.graphicFrame.Transform = new Transform());
            transform.Extents = new Drawing.Extents()
            {
                Cx = width,
                Cy = height
            };

            return this;
        }

        public OpenXmlElement CloneForSlide()
        {
            return this.graphicFrame.CloneNode(true);
        }
    }
}
