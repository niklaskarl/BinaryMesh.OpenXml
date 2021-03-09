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
                case GroupShape groupShape:
                    visual = new OpenXmlGroupShapeVisual(container, groupShape);
                    return true;
                case Picture picture:
                    visual = new OpenXmlPictureVisual(container, picture);
                    return true;
                case ConnectionShape connectionShape:
                    visual = new OpenXmlConnectionShapeVisual(container, connectionShape);
                    return true;
                default:
                    visual = null;
                    return false;
            }
        }

        private static bool TryCreateGraphicFrameVisual(IOpenXmlVisualContainer container, GraphicFrame graphicFrame, out IOpenXmlVisual visual)
        {
            string uri = graphicFrame.Graphic?.GraphicData?.Uri?.Value;
            switch (uri)
            {
                case "http://schemas.openxmlformats.org/drawingml/2006/chart":
                    visual = new OpenXmlChartVisual(container, graphicFrame);
                    return true;
                case "http://schemas.openxmlformats.org/drawingml/2006/table":
                    visual = new OpenXmlTableVisual(container, graphicFrame);
                    return true;
                default:
                    visual = new OpenXmlGraphicFrameVisual(container, graphicFrame);
                    return true;
            }
        }
    }
}
