using System;

namespace BinaryMesh.OpenXml.Presentations
{
    public interface ISlide
    {
        int Index { get; }

        ISlideLayout SlideLayout { get; }

        IShapeTree ShapeTree { get; }
    }
}
