using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;

using BinaryMesh.OpenXml.Presentations.Internal.Mixins;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal sealed class OpenXmlPictureVisual : IOpenXmlShapeElement, IOpenXmlTransformElement, IOpenXmlVisual, IVisual
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

        public IVisualTransform<IVisual> Transform => new OpenXmlVisualTransform<OpenXmlPictureVisual, IVisual>(this);

        IVisualTransform<IVisual> IVisual.Transform => this.Transform;

        public OpenXmlElement GetTransform()
        {
            return this.picture.ShapeProperties?.Transform2D;
        }

        public OpenXmlElement GetOrCreateTransform()
        {
            if (this.picture.ShapeProperties == null)
            {
                this.picture.ShapeProperties = new ShapeProperties();
            }

            
            if (this.picture.ShapeProperties.Transform2D == null)
            {
                this.picture.ShapeProperties.Transform2D = new Drawing.Transform2D();
            }

            return this.picture.ShapeProperties.Transform2D;
        }

        public OpenXmlElement GetShapeProperties()
        {
            return this.picture.ShapeProperties;
        }

        public OpenXmlElement GetOrCreateShapeProperties()
        {
            if (this.picture.ShapeProperties == null)
            {
                this.picture.ShapeProperties = new ShapeProperties();
            }

            return this.picture.ShapeProperties;
        }

        public OpenXmlElement CloneForSlide()
        {
            return this.picture.CloneNode(true);
        }
    }
}
