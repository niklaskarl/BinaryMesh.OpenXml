using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal sealed class OpenXmlShapeVisual : IShapeVisual, IOpenXmlVisual, IVisual
    {
        private readonly IOpenXmlVisualContainer container;

        private readonly Shape shape;

        public OpenXmlShapeVisual(IOpenXmlVisualContainer container, Shape shape)
        {
            this.container = container;
            this.shape = shape;
        }

        public uint Id => this.shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id;

        public string Name => this.shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name;

        public bool IsPlaceholder => this.shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape != null;

        public IShapeVisual AsShapeVisual()
        {
            return this;
        }

        public IGraphicFrameVisual AsGraphicFrameVisual()
        {
            return null;
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

        public IShapeVisual SetText(string text)
        {
            if (this.shape.TextBody == null)
            {
                this.shape.TextBody = new TextBody();
            }
            else
            {
                this.shape.TextBody.RemoveAllChildren<Drawing.Paragraph>();
            }

            this.shape.TextBody.AppendChildFluent(
                new Drawing.Paragraph()
                {
    
                }
                .AppendChildFluent(
                    new Drawing.Run()
                    {
                        Text = new Drawing.Text() { Text = text }
                    }
                )
            );

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
