using System;
using System.Collections.Generic;

namespace BinaryMesh.OpenXml.Presentations
{
    public interface IChartVisual : IVisual
    {
        new IVisualTransform<IChartVisual> Transform { get; }

        IChartSpace ChartSpace { get; }
    }
}
