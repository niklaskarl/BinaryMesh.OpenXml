using System;

namespace BinaryMesh.OpenXml.Presentations
{
    public interface IGraphicFrameVisual : IVisual
    {
        new IGraphicFrameVisual SetOffset(long x, long y);

        new IGraphicFrameVisual SetExtents(long width, long height);

        IGraphicFrameVisual SetContent(IChartSpace chartSpace);
    }
}
