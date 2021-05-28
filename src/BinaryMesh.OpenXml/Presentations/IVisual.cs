using System;

using BinaryMesh.OpenXml.Styles;

namespace BinaryMesh.OpenXml.Presentations
{
    public interface IVisual
    {
        uint Id { get; }

        string Name { get; }
        
        IVisualTransform<IVisual> Transform { get; }
    }
}
