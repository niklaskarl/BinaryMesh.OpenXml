using System;

namespace BinaryMesh.OpenXml.Presentations
{
    public interface IShapeVisual : IVisual
    {
        IShapeVisual SetText(string text);
    }
}
