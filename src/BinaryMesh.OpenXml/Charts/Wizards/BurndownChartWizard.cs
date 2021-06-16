using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;

using BinaryMesh.OpenXml.Spreadsheets;

namespace BinaryMesh.OpenXml.Charts.Wizards
{
    public struct BurndownChartData
    {
        public BurndownChartData(IEnumerable<BurndownChartCategory> categories) :
            this(categories.ToImmutableArray())
        {
        }

        public BurndownChartData(params BurndownChartCategory[] categories) :
            this(categories.ToImmutableArray())
        {
        }

        public BurndownChartData(ImmutableArray<BurndownChartCategory> categories)
        {
            this.Categories = categories;
            this.TotalCallback = null;
        }

        public ImmutableArray<BurndownChartCategory> Categories { get; }

        public Action<IChartValue<IBarChartValue>> TotalCallback { get; }
    }

    public struct BurndownChartCategory
    {
        public BurndownChartCategory(string name, IEnumerable<BurndownChartSeries> series) :
            this(name, series.ToImmutableArray())
        {
        }

        public BurndownChartCategory(string name, params BurndownChartSeries[] series) :
            this(name, series.ToImmutableArray())
        {
        }

        public BurndownChartCategory(string name, ImmutableArray<BurndownChartSeries> series)
        {
            this.Name = name;
            this.Series = series;
        }

        public string Name { get; }

        public ImmutableArray<BurndownChartSeries> Series { get; }
    }

    public struct BurndownChartSeries
    {
        public BurndownChartSeries(string name, double value)
        {
            this.Name = name;
            this.Value = value;
            this.Callback = null;
        }
        public BurndownChartSeries(string name, double value, Action<IChartValue<IBarChartValue>> callback)
        {
            this.Name = name;
            this.Value = value;
            this.Callback = callback;
        }

        public string Name { get; }

        public double Value { get; }

        public Action<IChartValue<IBarChartValue>> Callback { get; }
    }

    public sealed class BurndownChartWizard
    {
        private readonly IChartSpace chartSpace;

        private readonly BurndownChartData data;

        private readonly CartesianAxes axes;

        private readonly IBarChart barChart;

        private readonly ILineChart lineChart;

        private BurndownChartWizard(IChartSpace chartSpace, BurndownChartData data, CartesianAxes axes, IBarChart barChart, ILineChart lineChart)
        {
            this.chartSpace = chartSpace;
            this.data = data;
            this.axes = axes;
            this.barChart = barChart;
            this.lineChart = lineChart;
        }

        public static BurndownChartWizard BuildBurndownChart(IChartSpace chartSpace, BurndownChartData data)
        {
            using (ISpreadsheetDocument spreadsheet = chartSpace.OpenSpreadsheetDocument())
            {
                IWorksheet seriesSheet = spreadsheet.Workbook.AppendWorksheet("Series");
                IWorksheet connectorSheet = spreadsheet.Workbook.AppendWorksheet("Connector");
                
                double total = data.Categories.Sum(c => c.Series.Sum(s => s.Value));

                seriesSheet.Cells[0, 1].SetValue("Offset");
                seriesSheet.Cells[0, 2].SetValue("Total");
                
                seriesSheet.Cells[1, 0].SetValue("Total");
                connectorSheet.Cells[1, 0].SetValue("Total");

                seriesSheet.Cells[1, 1].SetValue(0);
                seriesSheet.Cells[1, 2].SetValue(total);

                double offset = total;
                uint column = 2;
                uint seriesRow = 3;
                uint connectorRow = 1;
                foreach (BurndownChartCategory category in data.Categories)
                {

                    // set name of category
                    seriesSheet.Cells[column, 0].SetValue(category.Name);
                    connectorSheet.Cells[column, 0].SetValue(category.Name);

                    // set connector of category
                    connectorSheet.Cells[column - 1, connectorRow].SetValue(offset);
                    connectorSheet.Cells[column, connectorRow].SetValue(offset);

                    // set offset of category
                    offset -= category.Series.Sum(s => s.Value);
                    seriesSheet.Cells[column, 1].SetValue(offset);

                    uint seriesIdx = 0;
                    uint seriesCount = (uint)category.Series.Length;
                    foreach (BurndownChartSeries series in category.Series)
                    {
                        // set name of series
                        seriesSheet.Cells[0, seriesRow + seriesCount - seriesIdx - 1].SetValue(series.Name);

                        // set value of series
                        seriesSheet.Cells[column, seriesRow + seriesCount - seriesIdx - 1].SetValue(Math.Abs(series.Value));

                        ++seriesIdx;
                    }

                    seriesRow += (uint)category.Series.Length;

                    ++connectorRow;
                    ++column;
                }
                
                CartesianAxes axes = chartSpace.AppendCartesianAxes();

                ILineChart lineChart = chartSpace.InsertLineChart(axes)
                    .InitializeFromRange(connectorSheet.GetRange(0, 1, 0, connectorRow - 1), connectorSheet.GetRange(1, 0, column - 1, 0));

                IBarChart barChart = chartSpace.InsertBarChart(axes)
                    .InitializeFromRange(seriesSheet.GetRange(0, 1, 0, seriesRow - 1), seriesSheet.GetRange(1, 0, column - 1, 0))
                    .SetDirection(BarChartDirection.Column)
                    .SetGrouping(BarChartGrouping.Stacked)
                    .SetOverlap(1.0);

                // hide offset series
                barChart.Series[0].Style.SetNoFill();
                
                if (data.TotalCallback != null)
                {
                    data.TotalCallback(barChart.Series[1].Values[1]);
                }

                int valueIndex = 1;
                int seriesIndex = 2;
                foreach (BurndownChartCategory category in data.Categories)
                {
                    int seriesIdx = 0;
                    foreach (BurndownChartSeries series in category.Series)
                    {
                        if (series.Callback != null)
                        {
                            IBarChartSeries s = barChart.Series[seriesIndex + category.Series.Length - seriesIdx - 1];
                            series.Callback(s.Values[valueIndex]);
                        }

                        ++seriesIdx;
                    }

                    seriesIndex += category.Series.Length;
                    ++valueIndex;
                }

                axes.CategoryAxis
                    .Text.SetFontSize(6)
                    .MajorGridlines.Remove();

                axes.ValueAxis
                    .SetVisibility(false)
                    .MajorGridlines.Remove();

                return new BurndownChartWizard(chartSpace, data, axes, barChart, lineChart);
            }
        }

        public BurndownChartWizard ConfigureTotal(Action<IBarChartValue> config)
        {
            IBarChartValue value = this.barChart.Series[1].Values[0];
            config(value);

            return this;
        }

        public BurndownChartWizard ConfigureSeries(int category, int series, Action<IBarChartValue> config)
        {
            int seriesIndex = data.Categories.Take(category + 1).Sum(c => c.Series.Length) + 1;
            int valueIndex = category + 1;
            IBarChartValue value = this.barChart.Series[seriesIndex].Values[valueIndex];
            config(value);

            return this;
        }

        public BurndownChartWizard ConfigureConnector(Action<ILineChartValue> config)
        {
            foreach (ILineChartValue value in lineChart.Series.SelectMany(s => s.Values))
            {
                config(value);
            }

            return this;
        }
    }
}
