using System;
using Packaging = DocumentFormat.OpenXml.Packaging;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal interface IOpenXmlSlide : IOpenXmlVisualContainer, ISlide
    {
        Packaging.SlidePart SlidePart { get; }
    }
}
