using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal sealed class OpenXmlGraphicFrameVisual : IOpenXmlVisual, IGraphicFrameVisual, IVisual
    {
        private readonly IOpenXmlPresentation presentation;

        private readonly GraphicFrame graphicFrame;

        public OpenXmlGraphicFrameVisual(IOpenXmlPresentation presentation, GraphicFrame graphicFrame)
        {
            this.presentation = presentation;
            this.graphicFrame = graphicFrame;
        }

        public uint Id => this.graphicFrame.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Id;

        public string Name => this.graphicFrame.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Name;

        public bool IsPlaceholder => this.graphicFrame.NonVisualGraphicFrameProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape?.HasCustomPrompt ?? false;

        public IShapeVisual AsShapeVisual()
        {
            return null;
        }

        public IGraphicFrameVisual AsGraphicFrameVisual()
        {
            return this;
        }

        public IGraphicFrameVisual SetExtend(double width, double height)
        {
            throw new NotImplementedException();
        }

        public IGraphicFrameVisual SetOrigin(double x, double y)
        {
            throw new NotImplementedException();
        }

        public IGraphicFrameVisual SetContent(IChartSpace chartSpace)
        {
            throw new NotImplementedException();
        }

        public OpenXmlElement CloneForSlide()
        {
            return this.graphicFrame.CloneNode(true);
        }

        IVisual IVisual.SetOrigin(double x, double y)
        {
            return this.SetOrigin(x, y);
        }

        IVisual IVisual.SetExtend(double width, double height)
        {
            return this.SetExtend(width, height);
        }
    }
}
