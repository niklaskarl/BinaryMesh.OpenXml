using System;
using DocumentFormat.OpenXml.Drawing;

namespace BinaryMesh.OpenXml.Presentations
{
    public interface IStyleableShape<TSelf>
    {
        TSelf SetFill(OpenXmlColor color);

        TSelf SetStroke(OpenXmlColor color);
    }
}
