using System;
using DocumentFormat.OpenXml.Packaging;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal interface ISlideRef : ISlide
    {
        SlidePart SlidePart { get; }
    }
}
