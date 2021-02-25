using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;

using BinaryMesh.OpenXml.Helpers;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal sealed class OpenXmlSlide : IOpenXmlSlide, ISlide
    {
        private readonly IOpenXmlPresentation presentation;

        private readonly SlidePart slidePart;

        public OpenXmlSlide(IOpenXmlPresentation presentation, SlidePart slidePart)
        {
            this.presentation = presentation;
            this.slidePart = slidePart;
        }

        public SlidePart SlidePart => this.slidePart;

        public int Index => throw new NotImplementedException();

        public KeyedReadOnlyList<string, IVisual> VisualTree => new EnumerableKeyedList<IOpenXmlVisual, string, IVisual>(
            this.slidePart.Slide.CommonSlideData.ShapeTree.Select(element => OpenXmlVisualFactory.TryCreateVisual(this.presentation, element, out IOpenXmlVisual visual) ? visual : null).Where(visual => visual != null),
            visual => visual.Name,
            visual => visual
        );
    }
}
