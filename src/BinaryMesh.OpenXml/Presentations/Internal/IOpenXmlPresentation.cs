using System;
using Packaging = DocumentFormat.OpenXml.Packaging;

using BinaryMesh.OpenXml.Internal;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal interface IOpenXmlPresentation : IOpenXmlDocument,  IPresentation
    {
        Packaging.PresentationPart PresentationPart { get; }
    }
}
