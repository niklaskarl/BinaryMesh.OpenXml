using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal sealed class OpenXmlConnectionVisual : IConnectionVisual, IOpenXmlVisual, IVisual
    {
        private readonly IOpenXmlVisualContainer container;

        private readonly ConnectionShape connectionShape;

        public OpenXmlConnectionVisual(IOpenXmlVisualContainer container, ConnectionShape connectionShape)
        {
            this.container = container;
            this.connectionShape = connectionShape;
        }

        public uint Id => this.connectionShape.NonVisualConnectionShapeProperties?.NonVisualDrawingProperties?.Id;

        public string Name => this.connectionShape.NonVisualConnectionShapeProperties?.NonVisualDrawingProperties?.Name;

        public bool IsPlaceholder => this.connectionShape.NonVisualConnectionShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape != null;

        private OpenXmlElement GetShapeProperties()
        {
            return this.connectionShape.ShapeProperties;
        }

        private OpenXmlElement GetOrCreateShapeProperties()
        {
            if (this.connectionShape.ShapeProperties == null)
            {
                this.connectionShape.ShapeProperties = new ShapeProperties();
            }

            return this.connectionShape.ShapeProperties;
        }

        public IShapeVisual AsShapeVisual()
        {
            return null;
        }

        public IConnectionVisual SetOffset(long x, long y)
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

        public IConnectionVisual SetExtents(long width, long height)
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

        public IConnectionVisual SetStroke(OpenXmlColor color)
        {
            OpenXmlElement shapeProperties = this.GetOrCreateShapeProperties();
            Drawing.Outline outline = shapeProperties.GetFirstChild<Drawing.Outline>() ?? shapeProperties.AppendChild(new Drawing.Outline() { Width = 12700 });
            outline.RemoveAllChildren<Drawing.NoFill>();
            outline.RemoveAllChildren<Drawing.SolidFill>();
            outline.RemoveAllChildren<Drawing.GradientFill>();
            outline.RemoveAllChildren<Drawing.BlipFill>();
            outline.RemoveAllChildren<Drawing.PatternFill>();
            outline.RemoveAllChildren<Drawing.GroupFill>();

            outline.AppendChild(new Drawing.SolidFill().AppendChildFluent(color.CreateColorElement()));

            return this;
        }

        public OpenXmlElement CloneForSlide()
        {
            return this.connectionShape.CloneNode(true);
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
