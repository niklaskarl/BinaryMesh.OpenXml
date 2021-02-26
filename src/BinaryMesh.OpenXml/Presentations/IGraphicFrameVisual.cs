using System;

namespace BinaryMesh.OpenXml.Presentations
{
    public interface IGraphicFrameVisual : IVisual
    {
        new IGraphicFrameVisual SetOrigin(double x, double y);

        new IGraphicFrameVisual SetExtend(double width, double height);

        IGraphicFrameVisual SetContent(IChartSpace chartSpace);
    }
}
