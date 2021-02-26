using System;

namespace BinaryMesh.OpenXml.Presentations
{
    public interface ISlide
    {
        int Index { get; }

        KeyedReadOnlyList<string, IVisual> VisualTree { get; }

        IShapeVisual AppendShapeVisual(string name);

        IGraphicFrameVisual AppendGraphicFrameVisual(string name);

        IChartSpace CreateChartSpace();
    }
}
