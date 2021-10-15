using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;

using BinaryMesh.OpenXml.Styles;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal abstract class OpenXmlGraphicFrameVisual<TSelf> : IOpenXmlVisual, IVisualTransform<IVisual>, IVisual where TSelf : IVisual
    {
        protected readonly IOpenXmlVisualContainer container;

        protected readonly GraphicFrame graphicFrame;

        public OpenXmlGraphicFrameVisual(IOpenXmlVisualContainer container, GraphicFrame graphicFrame)
        {
            this.container = container;
            this.graphicFrame = graphicFrame;
        }

        protected abstract TSelf Self { get; }

        public IOpenXmlVisualContainer Container => this.container;

        public uint Id => this.graphicFrame.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Id;

        public string Name => this.graphicFrame.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Name;

        public bool IsPlaceholder => this.graphicFrame.NonVisualGraphicFrameProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape?.HasCustomPrompt ?? false;

        IVisualTransform<IVisual> IVisual.Transform => this;

        public TSelf SetExtents(OpenXmlSize size)
        {
            this.SetExtents(size);
            return this.Self;
        }

        public TSelf SetExtents(long width, long height)
        {
            Transform transform = this.graphicFrame.Transform ?? (this.graphicFrame.Transform = new Transform());
            transform.Extents = new Drawing.Extents()
            {
                Cx = width,
                Cy = height
            };

            return this.Self;
        }

        public TSelf SetOffset(OpenXmlPoint point)
        {
            this.SetOffset(point);
            return this.Self;
        }

        public TSelf SetOffset(long x, long y)
        {
            Transform transform = this.graphicFrame.Transform ?? (this.graphicFrame.Transform = new Transform());
            transform.Offset = new Drawing.Offset()
            {
                X = x,
                Y = y
            };

            return this.Self;
        }

        public TSelf SetRect(OpenXmlRect rect)
        {
            this.SetOffset(rect.Left, rect.Top);
            this.SetExtents(rect.Width, rect.Height);

            return this.Self;
        }

        IVisual IVisualTransform<IVisual>.SetOffset(long x, long y)
        {
            return this.SetOffset(x, y);
        }

        IVisual IVisualTransform<IVisual>.SetExtents(long width, long height)
        {
            return this.SetExtents(width, height);
        }

        IVisual IVisualTransform<IVisual>.SetOffset(OpenXmlPoint point)
        {
            return this.SetOffset(point.Left, point.Top);
        }

        IVisual IVisualTransform<IVisual>.SetExtents(OpenXmlSize size)
        {
            return this.SetExtents(size.Width, size.Height);
        }

        IVisual IVisualTransform<IVisual>.SetRect(OpenXmlRect rect)
        {
            return this.SetRect(rect);
        }

        public OpenXmlElement CloneForSlide()
        {
            return this.graphicFrame.CloneNode(true);
        }
    }
}
