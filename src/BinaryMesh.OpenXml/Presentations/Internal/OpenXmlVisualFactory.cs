using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal static class OpenXmlVisualFactory
    {
        public static bool TryCreateVisual(IOpenXmlVisualContainer container, OpenXmlElement element, out IOpenXmlVisual visual)
        {
            switch (element)
            {
                case Shape shape:
                    visual = new OpenXmlShapeVisual(container, shape);
                    return true;
                case GraphicFrame graphicFrame:
                    visual = new OpenXmlGraphicFrameVisual(container, graphicFrame);
                    return true;
                default:
                    visual = null;
                    return false;
            }
        }
    }
}
