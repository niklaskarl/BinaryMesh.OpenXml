using System;
using DocumentFormat.OpenXml.Packaging;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal interface IPresentationRef : IPresentation
    {
        PresentationPart PresentationPart { get; }
    }
}
