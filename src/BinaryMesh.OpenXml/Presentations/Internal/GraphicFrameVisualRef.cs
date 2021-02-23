using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal sealed class GraphicFrameVisualRef : IVisualRef, IVisual
    {
        private readonly IPresentationRef presentation;

        private readonly GraphicFrame graphicFrame;

        public GraphicFrameVisualRef(IPresentationRef presentation, GraphicFrame graphicFrame)
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

        public OpenXmlElement CloneForSlide()
        {
            return this.graphicFrame.CloneNode(true);
        }
    }
}
