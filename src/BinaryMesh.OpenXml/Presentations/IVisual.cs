using System;

namespace BinaryMesh.OpenXml.Presentations
{
    public interface IVisual
    {
        uint Id { get; }

        string Name { get; }

        IShapeVisual AsShapeVisual();

        IGraphicFrameVisual AsGraphicFrameVisual();

        IVisual SetOrigin(double x, double y);

        IVisual SetExtend(double width, double height);
    }
}
