using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;

using BinaryMesh.OpenXml.Styles.Internal.Mixins;

namespace BinaryMesh.OpenXml.Styles.Internal
{
    internal class OpenXmlVisualTransform<TElement, TFluent> : IVisualTransform<TFluent>
        where TElement : IOpenXmlTransformElement, TFluent
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
            OpenXmlElement transform = this.element.GetOrCreateTransform();
            Offset offset = transform.GetFirstChild<Offset>() ?? transform.AppendChild(new Offset());
            offset.X = x;
            offset.Y = y;

            return this.element;
        }

        public TFluent SetExtents(OpenXmlSize size)
        {
            return this.SetExtents(size.Width, size.Height);
        }

        public TFluent SetExtents(long width, long height)
        {
            OpenXmlElement transform = this.element.GetOrCreateTransform();
            Extents extents = transform.GetFirstChild<Extents>() ?? transform.AppendChild(new Extents());
            extents.Cx = width;
            extents.Cy = height;

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
