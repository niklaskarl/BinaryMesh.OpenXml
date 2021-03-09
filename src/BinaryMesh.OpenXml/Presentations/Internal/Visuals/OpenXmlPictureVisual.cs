using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal sealed class OpenXmlPictureVisual : IOpenXmlVisual, IVisual
    {
        private readonly IOpenXmlVisualContainer container;

        private readonly Picture picture;

        public OpenXmlPictureVisual(IOpenXmlVisualContainer container, Picture picture)
        {
            this.container = container;
            this.picture = picture;
        }

        public uint Id => this.picture.NonVisualPictureProperties?.NonVisualDrawingProperties?.Id;

        public string Name => this.picture.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name;

        public bool IsPlaceholder => this.picture.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape != null;

        public IShapeVisual AsShapeVisual()
        {
            return null;
        }

        public IVisual SetOffset(long x, long y)
        {
            ShapeProperties shapeProperties = this.picture.ShapeProperties ?? (this.picture.ShapeProperties = new ShapeProperties());
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
            ShapeProperties shapeProperties = this.picture.ShapeProperties ?? (this.picture.ShapeProperties = new ShapeProperties());
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
            return this.picture.CloneNode(true);
        }
    }
}
