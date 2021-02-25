using System;
using Packaging = DocumentFormat.OpenXml.Packaging;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal interface IOpenXmlSlideMaster : ISlideMaster
    {
        Packaging.SlideMasterPart SlideMasterPart { get; }
    }
}
