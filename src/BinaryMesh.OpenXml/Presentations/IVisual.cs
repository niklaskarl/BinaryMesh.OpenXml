using System;

namespace BinaryMesh.OpenXml.Presentations
{
    public interface IVisual
    {
        uint Id { get; }

        string Name { get; }

        IShapeVisual AsShapeVisual();

        IVisual SetOffset(long x, long y);

        IVisual SetExtents(long width, long height);
    }
}
