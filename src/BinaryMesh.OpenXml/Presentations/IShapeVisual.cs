using System;

namespace BinaryMesh.OpenXml.Presentations
{
    public interface IShapeVisual : IVisual
    {
        new IShapeVisual SetOrigin(double x, double y);

        new IShapeVisual SetExtend(double width, double height);

        IShapeVisual SetText(string text);
    }
}
