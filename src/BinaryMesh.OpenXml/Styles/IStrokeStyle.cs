using System;

namespace BinaryMesh.OpenXml.Styles
{
    public interface IStrokeStyle<out TFluent>
    {
        TFluent SetStroke(OpenXmlColor color);

        TFluent SetStrokeWidth(double pt);
    }
}
