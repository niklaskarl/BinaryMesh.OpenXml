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

        IBarChart SetGapWidth(double ratio);

        IBarChart SetOverlap(double ratio);

        IBarChart InitializeFromRange(IRange labelRange, IRange categoryRange);

        IReadOnlyList<IBarChartSeries> Series { get; }
    }

    public interface IDataLabel<out TFluent>
    {
        IVisualStyle<TFluent> Style { get; }

        ITextStyle<TFluent> Text { get; }

        TFluent SetShowValue(bool show);

        TFluent SetShowPercent(bool show);

        TFluent SetShowCategoryName(bool show);

        TFluent SetShowLegendKey(bool show);

        TFluent SetShowSeriesName(bool show);

        TFluent SetShowBubbleSize(bool show);

        TFluent SetShowLeaderLines(bool show);

        TFluent Clear();
    }

    public interface IBarChartSeries : IChartSeries<IBarChartSeries, IBarChartValue>
    {
    }

    public interface IBarChartValue : IChartValue<IBarChartValue>
    {
    }
}
