using System;
using DocumentFormat.OpenXml;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal interface IVisualRef : IVisual
    {
        bool IsPlaceholder { get; }

        OpenXmlElement CloneForSlide();
    }
}
