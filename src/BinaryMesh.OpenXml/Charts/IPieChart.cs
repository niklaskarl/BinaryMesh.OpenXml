using System;
using BinaryMesh.OpenXml.Spreadsheets;

namespace BinaryMesh.OpenXml.Charts
{
    public interface IPieChart : IChart
    {
        IPieChart SetFirstSliceAngle(double rad);

        IPieChart SetExplosion(double percent);

        IPieChart SetHoleSize(double percent);

        IPieChartSeries Series { get; }
    }

    public interface IPieChartSeries : IChartSeries<IPieChartSeries, IPieChartValue>
    {
        IPieChartSeries SetText(IRange range);

        IPieChartSeries SetCategoryAxis(IRange range);

        IPieChartSeries SetValueAxis(IRange range);
    }

    public interface IPieChartValue : IChartValue<IPieChartValue>
    {
    }
}
