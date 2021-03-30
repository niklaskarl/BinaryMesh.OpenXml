using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;

using BinaryMesh.OpenXml.Presentations.Internal.Mixins;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal sealed class OpenXmlConnectionVisual : IOpenXmlShapeElement, IOpenXmlTransformElement, IOpenXmlVisual, IConnectionVisual, IVisual
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

        public IVisualStyle<IConnectionVisual> Style => new OpenXmlVisualStyle<OpenXmlConnectionVisual, IConnectionVisual>(this);

        public IVisualTransform<IConnectionVisual> Transform => new OpenXmlVisualTransform<OpenXmlConnectionVisual, IConnectionVisual>(this);

        IVisualTransform<IVisual> IVisual.Transform => this.Transform;

        public OpenXmlElement GetTransform()
        {
            return this.connectionShape.ShapeProperties?.Transform2D;
        }

        public OpenXmlElement GetOrCreateTransform()
        {
            if (this.connectionShape.ShapeProperties == null)
            {
                this.connectionShape.ShapeProperties = new ShapeProperties();
            }

            
            if (this.connectionShape.ShapeProperties.Transform2D == null)
            {
                this.connectionShape.ShapeProperties.Transform2D = new Drawing.Transform2D();
            }

            return this.connectionShape.ShapeProperties.Transform2D;
        }

        public OpenXmlElement GetShapeProperties()
        {
            return this.connectionShape.ShapeProperties;
        }

        public OpenXmlElement GetOrCreateShapeProperties()
        {
            if (this.connectionShape.ShapeProperties == null)
            {
                this.connectionShape.ShapeProperties = new ShapeProperties();
            }

            return this.connectionShape.ShapeProperties;
        }

        public OpenXmlElement CloneForSlide()
        {
            return this.connectionShape.CloneNode(true);
        }
    }
}
