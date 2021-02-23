using System;
using System.Collections.Generic;
using System.IO;

namespace BinaryMesh.OpenXml.Presentations
{
    public interface IPresentation : IDisposable
    {
        IReadOnlyList<ISlideMaster> SlideMasters { get; }

        IReadOnlyList<ISlide> Slides { get; }

        ISlide InsertSlide(ISlideLayout layout);

        ISlide InsertSlide(ISlideLayout layout, int index);

        IChart CreateChart();

        void Close(Stream destination);
    }
}
