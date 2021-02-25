using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

using BinaryMesh.OpenXml.Helpers;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal sealed class OpenXmlSlideMaster : IOpenXmlSlideMaster, ISlideMaster
    {
        private readonly IOpenXmlPresentation presentation;

        private readonly SlideMasterPart slideMasterPart;

        public OpenXmlSlideMaster(IOpenXmlPresentation presentation, SlideMasterPart slideMasterPart)
        {
            this.presentation = presentation;
            this.slideMasterPart = slideMasterPart;
        }

        public SlideMasterPart SlideMasterPart => this.slideMasterPart;

        public IReadOnlyList<ISlideLayout> SlideLayouts => new EnumerableList<SlideLayoutId, ISlideLayout>(
            this.slideMasterPart.SlideMaster.SlideLayoutIdList.Elements<SlideLayoutId>(),
            sl => new OpenXmlSlideLayout(this.presentation, this.slideMasterPart.GetPartById(sl.RelationshipId) as SlideLayoutPart)
        );
    }
}
