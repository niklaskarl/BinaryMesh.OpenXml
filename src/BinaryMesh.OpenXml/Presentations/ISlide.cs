using System;

namespace BinaryMesh.OpenXml.Presentations
{
    public interface ISlide
    {
        int Index { get; }

        KeyedReadOnlyList<string, IVisual> VisualTree { get; }

        IChartSpace CreateChartSpace();
    }
}
