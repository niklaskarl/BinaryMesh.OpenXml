using System;

namespace BinaryMesh.OpenXml.Styles
{
    public interface IStrokeStyle<out TFluent>
    {
        TFluent SetNoStroke();

        TFluent SetStroke(OpenXmlColor color);

        TFluent SetStrokeWidth(double pt);

        TFluent RemoveStrokeDash();

        TFluent SetStrokeDash(DocumentFormat.OpenXml.Drawing.PresetLineDashValues value);
    }
}
