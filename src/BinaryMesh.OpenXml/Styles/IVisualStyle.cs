using System;

namespace BinaryMesh.OpenXml.Styles
{
    public interface IVisualStyle<out TFluent> : IStrokeStyle<TFluent>
    {
        TFluent SetNoFill();

        TFluent SetFill(OpenXmlColor color);

        TFluent SetPresetGeometry(OpenXmlPresetGeometry presetGeometry);
    }
}
