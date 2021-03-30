using System;

namespace BinaryMesh.OpenXml.Presentations
{
    public interface IConnectionVisual : IVisual
    {
        new IVisualTransform<IConnectionVisual> Transform { get; }

        IVisualStyle<IConnectionVisual> Style { get; }
    }
}
