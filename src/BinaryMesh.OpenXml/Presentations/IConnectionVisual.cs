using System;

namespace BinaryMesh.OpenXml.Presentations
{
    public interface IConnectionVisual : IVisual
    {
        new IConnectionVisual SetOffset(long x, long y);

        new IConnectionVisual SetExtents(long width, long height);

        IConnectionVisual SetStroke(OpenXmlColor color);
    }
}
