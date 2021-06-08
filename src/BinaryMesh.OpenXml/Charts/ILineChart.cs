using System;
using System.Collections.Generic;
using BinaryMesh.OpenXml.Spreadsheets;

namespace BinaryMesh.OpenXml.Charts
{
    public interface ILineChart : IChart
    {
        ILineChart InitializeFromRange(IRange labelRange, IRange categoryRange);

        IReadOnlyList<ILineChartSeries> Series { get; }
    }

    public interface ILineChartSeries : IChartSeries<ILineChartSeries, ILineChartValue>
    {
    }

    public interface ILineChartValue : IChartValue<ILineChartValue>
    {
    }
}
