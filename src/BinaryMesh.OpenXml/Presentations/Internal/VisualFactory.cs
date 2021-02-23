using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal static class VisualFactory
    {
        public static bool TryCreateVisual(IPresentationRef presentation, OpenXmlElement element, out IVisualRef visual)
        {
            switch (element)
            {
                case Shape shape:
                    visual = new ShapeVisualRef(presentation, shape);
                    return true;
                case GraphicFrame graphicFrame:
                    visual = new GraphicFrameVisualRef(presentation, graphicFrame);
                    return true;
                default:
                    visual = null;
                    return false;
            }
        }
    }
}
