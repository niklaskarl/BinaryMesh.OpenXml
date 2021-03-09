using System;
using System.Collections.Generic;
using BinaryMesh.OpenXml.Spreadsheets;

namespace BinaryMesh.OpenXml.Presentations
{
    public enum BarChartDirection
    {
        Bar = DocumentFormat.OpenXml.Drawing.Charts.BarDirectionValues.Bar,
        Column = DocumentFormat.OpenXml.Drawing.Charts.BarDirectionValues.Column
    }

    public enum BarChartGrouping
    {
        PercentStacked = DocumentFormat.OpenXml.Drawing.Charts.BarGroupingValues.PercentStacked,
        Clustered = DocumentFormat.OpenXml.Drawing.Charts.BarGroupingValues.Clustered,
        Standard = DocumentFormat.OpenXml.Drawing.Charts.BarGroupingValues.Standard,
        Stacked = DocumentFormat.OpenXml.Drawing.Charts.BarGroupingValues.Stacked
    }

    public interface IBarChart : IChart
    {
        IBarChart SetDirection(BarChartDirection direction);

        IBarChart SetGrouping(BarChartGrouping grouping);

        IBarChart InitializeFromRange(IRange labelRange, IRange categoryRange);

        IReadOnlyList<IChartSeries> Series { get; }
    }
}
