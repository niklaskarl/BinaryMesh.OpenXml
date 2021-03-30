using System;

namespace BinaryMesh.OpenXml.Presentations
{
    public interface IVisual
    {
        uint Id { get; }

        string Name { get; }
        
        IVisualTransform<IVisual> Transform { get; }
    }
}
