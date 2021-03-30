using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;

using BinaryMesh.OpenXml.Presentations.Internal.Mixins;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal sealed class OpenXmlShapeVisual : IOpenXmlTextElement, IOpenXmlShapeElement, IOpenXmlTransformElement, IOpenXmlVisual, IShapeVisual, IVisual
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

        public IVisualStyle<IShapeVisual> Style => new OpenXmlVisualStyle<OpenXmlShapeVisual, IShapeVisual>(this);

        public ITextContent<IShapeVisual> Text => new OpenXmlTextContent<OpenXmlShapeVisual, IShapeVisual>(this);

        public IVisualTransform<IShapeVisual> Transform => new OpenXmlVisualTransform<OpenXmlShapeVisual, IShapeVisual>(this);

        IVisualTransform<IVisual> IVisual.Transform => this.Transform;

        public OpenXmlElement GetTextBody()
        {
            return this.shape.TextBody;
        }

        public OpenXmlElement GetOrCreateTextBody()
        {
            if (this.shape.TextBody == null)
            {
                this.shape.TextBody = new TextBody();
            }

            return this.shape.TextBody;
        }

        public OpenXmlElement GetTransform()
        {
            return this.shape.ShapeProperties?.Transform2D;
        }

        public OpenXmlElement GetOrCreateTransform()
        {
            if (this.shape.ShapeProperties == null)
            {
                this.shape.ShapeProperties = new ShapeProperties();
            }

            
            if (this.shape.ShapeProperties.Transform2D == null)
            {
                this.shape.ShapeProperties.Transform2D = new Drawing.Transform2D();
            }

            return this.shape.ShapeProperties.Transform2D;
        }

        public OpenXmlElement GetShapeProperties()
        {
            return this.shape.ShapeProperties;
        }

        public OpenXmlElement GetOrCreateShapeProperties()
        {
            if (this.shape.ShapeProperties == null)
            {
                this.shape.ShapeProperties = new ShapeProperties();
            }

            return this.shape.ShapeProperties;
        }

        public OpenXmlElement CloneForSlide()
        {
            return this.shape.CloneNode(true);
        }
    }
}
