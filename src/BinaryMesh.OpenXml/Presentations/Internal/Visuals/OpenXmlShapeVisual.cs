using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal sealed class OpenXmlShapeVisual : OpenXmlTextShapeBase<IShapeVisual>, IShapeVisual, IOpenXmlVisual, IVisual
    {
        private readonly IOpenXmlVisualContainer container;

        private readonly Shape shape;

        public OpenXmlShapeVisual(IOpenXmlVisualContainer container, Shape shape)
        {
            this.container = container;
            this.shape = shape;
        }

        protected override IShapeVisual Self => this;

        public uint Id => this.shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id;

        public string Name => this.shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name;

        public bool IsPlaceholder => this.shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape != null;

        protected override OpenXmlElement GetTextBody()
        {
            return this.shape.TextBody;
        }

        protected override OpenXmlElement GetOrCreateTextBody()
        {
            if (this.shape.TextBody == null)
            {
                this.shape.TextBody = new TextBody();
            }

            return this.shape.TextBody;
        }

        protected override OpenXmlElement GetShapeProperties()
        {
            return this.shape.ShapeProperties;
        }

        protected override OpenXmlElement GetOrCreateShapeProperties()
        {
            if (this.shape.ShapeProperties == null)
            {
                this.shape.ShapeProperties = new ShapeProperties();
            }

            return this.shape.ShapeProperties;
        }

        public IShapeVisual AsShapeVisual()
        {
            return this;
        }

        public IShapeVisual SetOffset(long x, long y)
        {
            ShapeProperties shapeProperties = this.shape.ShapeProperties ?? (this.shape.ShapeProperties = new ShapeProperties());
            Drawing.Transform2D transform = shapeProperties.Transform2D ?? (shapeProperties.Transform2D = new Drawing.Transform2D());
            transform.Offset = new Drawing.Offset()
            {
                X = x,
                Y = y
            };

            return this;
        }

        public IShapeVisual SetExtents(long width, long height)
        {
            ShapeProperties shapeProperties = this.shape.ShapeProperties ?? (this.shape.ShapeProperties = new ShapeProperties());
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
            return this.shape.CloneNode(true);
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
