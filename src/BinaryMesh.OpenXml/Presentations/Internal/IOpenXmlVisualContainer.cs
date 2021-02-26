using System;
using DocumentFormat.OpenXml.Packaging;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal interface IOpenXmlVisualContainer
    {
        OpenXmlPart Part { get; }
    }
}
