using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

using BinaryMesh.OpenXml.Helpers;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal sealed class SlideRef : ISlideRef, ISlide
    {
        private readonly IPresentationRef presentation;
        private readonly SlidePart slidePart;

        public SlideRef(IPresentationRef presentation, SlidePart slidePart)
        {
            this.presentation = presentation;
            this.slidePart = slidePart;
        }

        public SlidePart SlidePart => this.slidePart;

        public int Index => throw new NotImplementedException();

        public KeyedReadOnlyList<string, IVisual> VisualTree => new EnumerableKeyedList<IVisualRef, string, IVisual>(
            this.slidePart.Slide.CommonSlideData.ShapeTree.Select(element => VisualFactory.TryCreateVisual(this.presentation, element, out IVisualRef visual) ? visual : null).Where(visual => visual != null),
            visual => visual.Name,
            visual => visual
        );
    }
}
