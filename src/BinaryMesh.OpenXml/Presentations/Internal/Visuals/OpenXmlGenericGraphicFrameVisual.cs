using System;
using DocumentFormat.OpenXml.Presentation;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal sealed class OpenXmlGenericGraphicFrameVisual : OpenXmlGraphicFrameVisual<IVisual>, IVisual
    {
        public OpenXmlGenericGraphicFrameVisual(IOpenXmlVisualContainer container, GraphicFrame graphicFrame) :
            base(container, graphicFrame)
        {
        }

        protected override IVisual Self => this;
    }
}
