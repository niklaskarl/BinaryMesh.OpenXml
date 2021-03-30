using System;

namespace BinaryMesh.OpenXml.Presentations
{
    public interface IVisualTransform<out TFluent>
    {
        TFluent SetOffset(OpenXmlPoint point);

        TFluent SetOffset(long x, long y);

        TFluent SetExtents(OpenXmlSize size);

        TFluent SetExtents(long width, long height);

        TFluent SetRect(OpenXmlRect rect);
    }
}
