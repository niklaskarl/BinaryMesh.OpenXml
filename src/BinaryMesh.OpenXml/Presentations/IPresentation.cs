using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace BinaryMesh.OpenXml.Presentations
{
    public interface IPresentation : IDisposable
    {
        IReadOnlyList<ISlideMaster> SlideMasters { get; }

        IReadOnlyList<ISlide> Slides { get; }

        ISlide InsertSlide(ISlideLayout layout);

        ISlide InsertSlide(ISlideLayout layout, int index);

        void Close(Stream destination);

        Task CloseAsync(Stream destination);
    }
}
