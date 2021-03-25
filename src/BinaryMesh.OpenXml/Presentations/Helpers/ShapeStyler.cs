using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;

namespace BinaryMesh.OpenXml.Presentations.Helpers
{
    internal static class OpenXmlShapeStyler
    {
        public static void SetFill(OpenXmlElement shapeProperties, OpenXmlColor color)
        {
            shapeProperties.RemoveAllChildren<NoFill>();
            shapeProperties.RemoveAllChildren<SolidFill>();
            shapeProperties.RemoveAllChildren<GradientFill>();
            shapeProperties.RemoveAllChildren<BlipFill>();
            shapeProperties.RemoveAllChildren<PatternFill>();
            shapeProperties.RemoveAllChildren<GroupFill>();

            shapeProperties.AppendChild(new SolidFill().AppendChildFluent(color.CreateColorElement()));
        }

        public static void SetStroke(OpenXmlElement shapeProperties, OpenXmlColor color)
        {
            Outline outline = shapeProperties.GetFirstChild<Outline>() ?? shapeProperties.AppendChild(new Outline() { Width = 12700 });
            outline.RemoveAllChildren<NoFill>();
            outline.RemoveAllChildren<SolidFill>();
            outline.RemoveAllChildren<GradientFill>();
            outline.RemoveAllChildren<BlipFill>();
            outline.RemoveAllChildren<PatternFill>();
            outline.RemoveAllChildren<GroupFill>();

            outline.AppendChild(new SolidFill().AppendChildFluent(color.CreateColorElement()));
        }
    }
}
