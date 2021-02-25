using System;
using Packaging = DocumentFormat.OpenXml.Packaging;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal interface IOpenXmlPresentation : IPresentation
    {
        Packaging.PresentationPart PresentationPart { get; }
    }
}
