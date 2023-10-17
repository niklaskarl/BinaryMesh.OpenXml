using System;
using System.Collections.Generic;
using BinaryMesh.OpenXml.Spreadsheets;

namespace BinaryMesh.OpenXml.Charts
{
    public enum LineChartGrouping
    {
        PercentStacked = DocumentFormat.OpenXml.Drawing.Charts.GroupingValues.PercentStacked,
        Standard = DocumentFormat.OpenXml.Drawing.Charts.GroupingValues.Standard,
        Stacked = DocumentFormat.OpenXml.Drawing.Charts.GroupingValues.Stacked
    }

    public interface ILineChart : IChart
    {
        ILineChart SetGrouping(LineChartGrouping grouping);

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
