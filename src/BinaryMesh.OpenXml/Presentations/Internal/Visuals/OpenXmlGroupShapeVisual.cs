using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal sealed class OpenXmlGroupShapeVisual : IOpenXmlVisual, IVisual
    {
        private readonly IOpenXmlVisualContainer container;

        private readonly GroupShape groupShape;

        public OpenXmlGroupShapeVisual(IOpenXmlVisualContainer container, GroupShape groupShape)
        {
            this.container = container;
            this.groupShape = groupShape;
        }

        public uint Id => this.groupShape.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?.Id;

        public string Name => this.groupShape.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?.Name;

        public bool IsPlaceholder => this.groupShape.NonVisualGroupShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape != null;

        public IShapeVisual AsShapeVisual()
        {
            return null;
        }

        public IVisual SetOffset(long x, long y)
        {
            GroupShapeProperties groupShapeProperties = this.groupShape.GroupShapeProperties ?? (this.groupShape.GroupShapeProperties = new GroupShapeProperties());
            Drawing.TransformGroup transform = groupShapeProperties.TransformGroup ?? (groupShapeProperties.TransformGroup = new Drawing.TransformGroup());
            transform.Offset = new Drawing.Offset()
            {
                X = x,
                Y = y
            };

            return this;
        }

        public IVisual SetExtents(long width, long height)
        {
            GroupShapeProperties groupShapeProperties = this.groupShape.GroupShapeProperties ?? (this.groupShape.GroupShapeProperties = new GroupShapeProperties());
            Drawing.TransformGroup transform = groupShapeProperties.TransformGroup ?? (groupShapeProperties.TransformGroup = new Drawing.TransformGroup());
            transform.Extents = new Drawing.Extents()
            {
                Cx = width,
                Cy = height
            };

            return this;
        }

        public OpenXmlElement CloneForSlide()
        {
            return this.groupShape.CloneNode(true);
        }
    }
}
