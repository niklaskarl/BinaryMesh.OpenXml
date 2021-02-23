using System;
using System.Collections.Generic;

namespace BinaryMesh.OpenXml.Presentations
{
    public interface ISlideMaster
    {
        IReadOnlyList<ISlideLayout> SlideLayouts { get; }
    }
}
