using System;
using DocumentFormat.OpenXml.Packaging;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal sealed class SlideLayoutRef : ISlideLayoutRef, ISlideLayout
    {
        private readonly IPresentationRef presentation;

        private readonly SlideLayoutPart slideLayoutPart;

        public SlideLayoutRef(IPresentationRef presentation, SlideLayoutPart slideLayoutPart)
        {
            this.presentation = presentation;
            this.slideLayoutPart = slideLayoutPart;
        }

        public SlideLayoutPart SlideLayoutPart => this.slideLayoutPart;
    }
}
