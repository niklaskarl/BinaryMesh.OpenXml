using System;

namespace BinaryMesh.OpenXml.Presentations
{
    public interface IShapeVisual : ITextShape<IShapeVisual>, IVisual
    {
        new IShapeVisual SetOffset(long x, long y);

        new IShapeVisual SetExtents(long width, long height);
    }
}
