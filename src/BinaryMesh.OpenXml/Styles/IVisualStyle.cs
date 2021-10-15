using System;

namespace BinaryMesh.OpenXml.Styles
{
    public interface IFillStyle<out TFluent>
    {
        TFluent SetNoFill();

        TFluent SetFill(OpenXmlColor color);

        TFluent SetPatternFill(DocumentFormat.OpenXml.Drawing.PresetPatternValues pattern, OpenXmlColor background, OpenXmlColor foreground);
    }

    public interface IVisualStyle<out TFluent> : IFillStyle<TFluent>, IStrokeStyle<TFluent>
    {
        TFluent SetPresetGeometry(OpenXmlPresetGeometry presetGeometry);
    }
}
