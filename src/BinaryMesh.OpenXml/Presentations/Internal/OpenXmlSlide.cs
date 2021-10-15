using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;

using BinaryMesh.OpenXml.Internal;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal sealed class OpenXmlSlide : IOpenXmlSlide, IOpenXmlVisualContainer, ISlide
    {
        private readonly IOpenXmlPresentation presentation;

        private readonly SlidePart slidePart;

        public OpenXmlSlide(IOpenXmlPresentation presentation, SlidePart slidePart)
        {
            this.presentation = presentation;
            this.slidePart = slidePart;
        }

        public IOpenXmlDocument Document => this.presentation;

        public OpenXmlPart Part => this.slidePart;

        public SlidePart SlidePart => this.slidePart;

        public int Index => throw new NotImplementedException();

        public ISlideLayout SlideLayout => new OpenXmlSlideLayout(this.presentation, this.slidePart.GetPartsOfType<SlideLayoutPart>().Single());

        public IShapeTree ShapeTree => new OpenXmlShapeTree(this, this.slidePart.Slide.CommonSlideData.ShapeTree);
    }
}
