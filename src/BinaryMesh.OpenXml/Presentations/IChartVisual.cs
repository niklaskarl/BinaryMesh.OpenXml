using System;
using System.Collections.Generic;

namespace BinaryMesh.OpenXml.Presentations
{
    public interface IChartVisual : IVisual
    {
        new IChartVisual SetOffset(long x, long y);

        new IChartVisual SetExtents(long width, long height);

        IChartSpace ChartSpace { get; }
    }
}
