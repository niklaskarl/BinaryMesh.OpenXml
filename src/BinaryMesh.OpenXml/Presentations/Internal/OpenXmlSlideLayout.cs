using System;
using DocumentFormat.OpenXml.Packaging;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal sealed class OpenXmlSlideLayout : IOpenXmlSlideLayout, ISlideLayout
    {
        private readonly IOpenXmlPresentation presentation;

        private readonly SlideLayoutPart slideLayoutPart;

        public OpenXmlSlideLayout(IOpenXmlPresentation presentation, SlideLayoutPart slideLayoutPart)
        {
            this.presentation = presentation;
            this.slideLayoutPart = slideLayoutPart;
        }

        public SlideLayoutPart SlideLayoutPart => this.slideLayoutPart;
    }
}
