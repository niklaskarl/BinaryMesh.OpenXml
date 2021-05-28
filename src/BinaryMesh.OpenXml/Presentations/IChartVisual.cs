using System;

using BinaryMesh.OpenXml.Charts;
using BinaryMesh.OpenXml.Styles;

namespace BinaryMesh.OpenXml.Presentations
{
    public interface IChartVisual : IVisual
    {
        new IVisualTransform<IChartVisual> Transform { get; }

        IChartSpace ChartSpace { get; }
    }
}
