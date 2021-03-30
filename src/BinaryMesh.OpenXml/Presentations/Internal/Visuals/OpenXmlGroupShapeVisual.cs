using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;

using BinaryMesh.OpenXml.Presentations.Internal.Mixins;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal sealed class OpenXmlGroupShapeVisual : IOpenXmlShapeElement, IOpenXmlTransformElement, IOpenXmlVisual, IVisual
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

        public IVisualTransform<IVisual> Transform => new OpenXmlVisualTransform<OpenXmlGroupShapeVisual, IVisual>(this);

        IVisualTransform<IVisual> IVisual.Transform => this.Transform;

        public OpenXmlElement GetTransform()
        {
            return this.groupShape.GroupShapeProperties?.TransformGroup;
        }

        public OpenXmlElement GetOrCreateTransform()
        {
            if (this.groupShape.GroupShapeProperties == null)
            {
                this.groupShape.GroupShapeProperties = new GroupShapeProperties();
            }

            
            if (this.groupShape.GroupShapeProperties.TransformGroup == null)
            {
                this.groupShape.GroupShapeProperties.TransformGroup = new Drawing.TransformGroup();
            }

            return this.groupShape.GroupShapeProperties.TransformGroup;
        }

        public OpenXmlElement GetShapeProperties()
        {
            return this.groupShape.GroupShapeProperties;
        }

        public OpenXmlElement GetOrCreateShapeProperties()
        {
            if (this.groupShape.GroupShapeProperties == null)
            {
                this.groupShape.GroupShapeProperties = new GroupShapeProperties();
            }

            return this.groupShape.GroupShapeProperties;
        }

        public OpenXmlElement CloneForSlide()
        {
            return this.groupShape.CloneNode(true);
        }
    }
}
