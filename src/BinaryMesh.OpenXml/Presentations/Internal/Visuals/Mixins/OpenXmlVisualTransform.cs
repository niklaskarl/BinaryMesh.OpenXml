using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;

namespace BinaryMesh.OpenXml.Presentations.Internal.Mixins
{
    internal class OpenXmlVisualTransform<TElement, TFluent> : IVisualTransform<TFluent>
        where TElement : IOpenXmlShapeElement, TFluent
    {
        protected readonly TElement element;

        public OpenXmlVisualTransform(TElement element)
        {
            this.element = element;
        }

        public TFluent SetOffset(OpenXmlPoint point)
        {
            return this.SetOffset(point.Left, point.Top);
        }

        public TFluent SetOffset(long x, long y)
        {
            OpenXmlElement shapeProperties = this.element.GetOrCreateShapeProperties();
            Transform2D transform = shapeProperties.GetFirstChild<Transform2D>() ?? shapeProperties.AppendChild(new Transform2D());
            transform.Offset = new Offset()
            {
                X = x,
                Y = y
            };

            return this.element;
        }

        public TFluent SetExtents(OpenXmlSize size)
        {
            return this.SetExtents(size.Width, size.Height);
        }

        public TFluent SetExtents(long width, long height)
        {
            OpenXmlElement shapeProperties = this.element.GetOrCreateShapeProperties();
            Transform2D transform = shapeProperties.GetFirstChild<Transform2D>() ?? shapeProperties.AppendChild(new Transform2D());
            transform.Extents = new Extents()
            {
                Cx = width,
                Cy = height
            };

            return this.element;
        }

        public TFluent SetRect(OpenXmlRect rect)
        {
            this.SetOffset(rect.Left, rect.Top);
            this.SetExtents(rect.Width, rect.Height);

            return this.element;
        }
    }
}
