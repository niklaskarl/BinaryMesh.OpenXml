using System;
using DocumentFormat.OpenXml.Packaging;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal interface ISlideMasterRef : ISlideMaster
    {
        SlideMasterPart SlideMasterPart { get; }
    }
}
