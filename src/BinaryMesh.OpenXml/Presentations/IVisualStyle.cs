using System;

namespace BinaryMesh.OpenXml.Presentations
{
    public interface IVisualStyle<out TFluent>
    {
        TFluent SetNoFill();

        TFluent SetFill(OpenXmlColor color);

        TFluent SetStroke(OpenXmlColor color);

        TFluent SetPresetGeometry(OpenXmlPresetGeometry presetGeometry);
    }
}
