using System;

using BinaryMesh.OpenXml.Styles;

namespace BinaryMesh.OpenXml.Presentations
{
    public interface IShapeVisual : IVisual
    {
        new IVisualTransform<IShapeVisual> Transform { get; }

        IVisualStyle<IShapeVisual> Style { get; }

        ITextContent<IShapeVisual> Text { get; }
    }
}
