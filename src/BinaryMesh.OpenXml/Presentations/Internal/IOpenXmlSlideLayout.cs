using System;
using Packaging = DocumentFormat.OpenXml.Packaging;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal interface IOpenXmlSlideLayout : IOpenXmlVisualContainer, ISlideLayout
    {
        Packaging.SlideLayoutPart SlideLayoutPart { get; }
    }
}
