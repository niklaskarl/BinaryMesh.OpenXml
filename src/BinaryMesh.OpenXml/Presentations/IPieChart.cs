using System;

namespace BinaryMesh.OpenXml.Presentations
{
    public interface IPieChart : IChart
    {
        IChartSeries Series { get; }
    }
}
