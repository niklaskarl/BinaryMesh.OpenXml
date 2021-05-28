using System;

using BinaryMesh.OpenXml.Styles;

namespace BinaryMesh.OpenXml.Presentations
{
    public interface IConnectionVisual : IVisual
    {
        new IVisualTransform<IConnectionVisual> Transform { get; }

        IStrokeStyle<IConnectionVisual> Style { get; }
    }
}
