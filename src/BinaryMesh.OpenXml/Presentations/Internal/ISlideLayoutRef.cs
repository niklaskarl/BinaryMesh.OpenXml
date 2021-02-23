using System;
using DocumentFormat.OpenXml.Packaging;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal interface ISlideLayoutRef : ISlideLayout
    {
        SlideLayoutPart SlideLayoutPart { get; }
    }
}
