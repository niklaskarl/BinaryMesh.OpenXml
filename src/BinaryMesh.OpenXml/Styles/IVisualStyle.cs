using System;

namespace BinaryMesh.OpenXml.Styles
{
    public interface IFillStyle<out TFluent>
    {
        TFluent SetNoFill();

        TFluent SetFill(OpenXmlColor color);
    }

    public interface IVisualStyle<out TFluent> : IFillStyle<TFluent>, IStrokeStyle<TFluent>
    {
        TFluent SetPresetGeometry(OpenXmlPresetGeometry presetGeometry);
    }
}
