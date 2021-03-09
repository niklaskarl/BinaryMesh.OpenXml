using System;

namespace BinaryMesh.OpenXml.Presentations
{
    public interface IPieChart : IChart
    {
        IPieChart SetFirstSliceAngle(double rad);

        IPieChart SetExplosion(double percent);

        IPieChart SetHoleSize(double percent);

        IChartSeries Series { get; }
    }
}
