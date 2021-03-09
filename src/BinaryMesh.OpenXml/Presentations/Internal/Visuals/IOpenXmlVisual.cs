using System;
using DocumentFormat.OpenXml;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal interface IOpenXmlVisual : IVisual
    {
        bool IsPlaceholder { get; }

        OpenXmlElement CloneForSlide();
    }
}
