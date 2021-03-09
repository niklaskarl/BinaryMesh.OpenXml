using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal sealed class OpenXmlConnectionShapeVisual : IOpenXmlVisual, IVisual
    {
        private readonly IOpenXmlVisualContainer container;

        private readonly ConnectionShape connectionShape;

        public OpenXmlConnectionShapeVisual(IOpenXmlVisualContainer container, ConnectionShape connectionShape)
        {
            this.container = container;
            this.connectionShape = connectionShape;
        }

        public uint Id => this.connectionShape.NonVisualConnectionShapeProperties?.NonVisualDrawingProperties?.Id;

        public string Name => this.connectionShape.NonVisualConnectionShapeProperties?.NonVisualDrawingProperties?.Name;

        public bool IsPlaceholder => this.connectionShape.NonVisualConnectionShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape != null;

        public IShapeVisual AsShapeVisual()
        {
            return null;
        }

        public IVisual SetOffset(long x, long y)
        {
            ShapeProperties shapeProperties = this.connectionShape.ShapeProperties ?? (this.connectionShape.ShapeProperties = new ShapeProperties());
            Drawing.Transform2D transform = shapeProperties.Transform2D ?? (shapeProperties.Transform2D = new Drawing.Transform2D());
            transform.Offset = new Drawing.Offset()
            {
                X = x,
                Y = y
            };

            return this;
        }

        public IVisual SetExtents(long width, long height)
        {
            ShapeProperties shapeProperties = this.connectionShape.ShapeProperties ?? (this.connectionShape.ShapeProperties = new ShapeProperties());
            Drawing.Transform2D transform = shapeProperties.Transform2D ?? (shapeProperties.Transform2D = new Drawing.Transform2D());
            transform.Extents = new Drawing.Extents()
            {
                Cx = width,
                Cy = height
            };

            return this;
        }

        public OpenXmlElement CloneForSlide()
        {
            return this.connectionShape.CloneNode(true);
        }
    }
}
