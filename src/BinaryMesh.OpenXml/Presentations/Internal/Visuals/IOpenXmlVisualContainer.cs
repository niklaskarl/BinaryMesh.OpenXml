using System;
using DocumentFormat.OpenXml.Packaging;

using BinaryMesh.OpenXml.Internal;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal interface IOpenXmlVisualContainer
    {
        IOpenXmlDocument Document { get; }

        OpenXmlPart Part { get; }
    }
}
