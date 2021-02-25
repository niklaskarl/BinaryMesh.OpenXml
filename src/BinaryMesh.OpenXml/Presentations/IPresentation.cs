using System;
using System.Collections.Generic;
using System.IO;

using BinaryMesh.OpenXml.Spreadsheets;

namespace BinaryMesh.OpenXml.Presentations
{
    public interface IPresentation : IDisposable
    {
        IReadOnlyList<ISlideMaster> SlideMasters { get; }

        IReadOnlyList<ISlide> Slides { get; }

        ISlide InsertSlide(ISlideLayout layout);

        ISlide InsertSlide(ISlideLayout layout, int index);

        IChartSpace CreateChartSpace();

        void Close(Stream destination);
    }
}
