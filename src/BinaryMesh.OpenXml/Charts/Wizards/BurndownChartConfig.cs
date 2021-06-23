using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;

using BinaryMesh.OpenXml.Spreadsheets;

namespace BinaryMesh.OpenXml.Charts.Wizards
{
    public interface IBurndownChartConfig
    {
        IBurndownChartConfig WithTotalStyle(Action<IBarChartValue> style);

        IBurndownChartConfig WithConnectorStyle(Action<ILineChartValue> style);

        IBurndownChartCategoryConfig AddCategory(string name);

        void Apply(IChartSpace chartSpace);
    }

    public interface IBurndownChartCategoryConfig : IBurndownChartConfig
    {
        IBurndownChartCategoryConfig WithCustomOffset(double offset);

        IBurndownChartValueConfig AddValue(string name, double value);
    }

    public interface IBurndownChartValueConfig : IBurndownChartConfig
    {
        IBurndownChartValueConfig AddValue(string name, double value);

        IBurndownChartValueConfig WithStyle(Action<IBarChartValue> style);
    }

    internal abstract class BurndownChartConfigProxy : IBurndownChartConfig
    {
        protected readonly IBurndownChartConfig target;

        protected BurndownChartConfigProxy(IBurndownChartConfig target)
        {
            this.target = target;
        }

        public IBurndownChartCategoryConfig AddCategory(string name)
        {
            return this.target.AddCategory(name);
        }

        public void Apply(IChartSpace chartSpace)
        {
            this.target.Apply(chartSpace);
        }

        public IBurndownChartConfig WithTotalStyle(Action<IBarChartValue> style)
        {
            return this.target.WithTotalStyle(style);
        }

        public IBurndownChartConfig WithConnectorStyle(Action<ILineChartValue> style)
        {
            return this.target.WithConnectorStyle(style);
        }
    }

    public sealed class BurndownChartConfig : IBurndownChartConfig
    {
        private readonly List<BurndownChartCategoryConfig> categories;

        private Action<IBarChartValue> totalStyle;

        private Action<ILineChartValue> connectorStyle;

        public BurndownChartConfig()
        {
            this.categories = new List<BurndownChartCategoryConfig>();
        }

        public IBurndownChartCategoryConfig AddCategory(string name)
        {
            BurndownChartCategoryConfig result = new BurndownChartCategoryConfig(this, name);
            this.categories.Add(result);
            
            return result;
        }

        public void Apply(IChartSpace chartSpace)
        {
            using (ISpreadsheetDocument spreadsheet = chartSpace.OpenSpreadsheetDocument())
            {
                IWorksheet seriesSheet = spreadsheet.Workbook.AppendWorksheet("Series");
                IWorksheet connectorSheet = spreadsheet.Workbook.AppendWorksheet("Connector");
                
                double total = this.categories.Sum(c => c.CustomOffset.HasValue ? c.CustomOffset.Value : c.Values.Sum(s => s.Value));

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
                foreach (BurndownChartCategoryConfig category in this.categories)
                {

                    // set name of category
                    seriesSheet.Cells[column, 0].SetValue(category.Name);
                    connectorSheet.Cells[column, 0].SetValue(category.Name);

                    // set connector of category
                    connectorSheet.Cells[column - 1, connectorRow].SetValue(offset);
                    connectorSheet.Cells[column, connectorRow].SetValue(offset);

                    // set offset of category
                    seriesSheet.Cells[column, 1].SetValue(offset - category.Values.Sum(s => s.Value));

                    if (category.CustomOffset.HasValue)
                    {
                        offset -= category.CustomOffset.Value;
                    }
                    else
                    {
                        offset -= category.Values.Sum(s => s.Value);
                    }

                    uint seriesIdx = 0;
                    uint seriesCount = (uint)category.Values.Count;
                    foreach (BurndownChartValueConfig value in category.Values)
                    {
                        // set name of series
                        seriesSheet.Cells[0, seriesRow + seriesCount - seriesIdx - 1].SetValue(value.Name);

                        // set value of series
                        seriesSheet.Cells[column, seriesRow + seriesCount - seriesIdx - 1].SetValue(Math.Abs(value.Value));

                        ++seriesIdx;
                    }

                    seriesRow += (uint)category.Values.Count;

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
                
                if (this.totalStyle != null)
                {
                    this.totalStyle(barChart.Series[1].Values[1]);
                }

                if (this.connectorStyle != null)
                {
                    foreach (ILineChartValue value in lineChart.Series.SelectMany(s => s.Values))
                    {
                        this.connectorStyle(value);
                    }
                }

                int valueIndex = 1;
                int seriesIndex = 2;
                foreach (BurndownChartCategoryConfig category in this.categories)
                {
                    int seriesIdx = 0;
                    foreach (BurndownChartValueConfig value in category.Values)
                    {
                        if (value.Style != null)
                        {
                            IBarChartSeries s = barChart.Series[seriesIndex + category.Values.Count - seriesIdx - 1];
                            value.Style(s.Values[valueIndex]);
                        }

                        ++seriesIdx;
                    }

                    seriesIndex += category.Values.Count;
                    ++valueIndex;
                }

                axes.CategoryAxis
                    .Text.SetFontSize(6)
                    .MajorGridlines.Remove();

                axes.ValueAxis
                    .SetVisibility(false)
                    .MajorGridlines.Remove();
            }
        }

        public IBurndownChartConfig WithConnectorStyle(Action<ILineChartValue> style)
        {
            this.connectorStyle = style;
            return this;
        }

        public IBurndownChartConfig WithTotalStyle(Action<IBarChartValue> style)
        {
            this.totalStyle = style;
            return this;
        }

        private sealed class BurndownChartCategoryConfig : BurndownChartConfigProxy, IBurndownChartCategoryConfig
        {
            private readonly BurndownChartConfig config;

            private readonly List<BurndownChartValueConfig> values;

            private string name;

            private double? customOffset;

            public BurndownChartCategoryConfig(BurndownChartConfig config, string name) :
                base(config)
            {
                this.config = config;
                this.values = new List<BurndownChartValueConfig>();
                this.name = name;
            }

            internal List<BurndownChartValueConfig> Values => values;

            internal string Name => this.name;

            internal double? CustomOffset => this.customOffset;

            public IBurndownChartValueConfig AddValue(string name, double value)
            {
                BurndownChartValueConfig result = new BurndownChartValueConfig(this, name, value);
                this.values.Add(result);
                
                return result;
            }

            public IBurndownChartCategoryConfig WithCustomOffset(double offset)
            {
                this.customOffset = offset;
                return this;
            }
        }

        private sealed class BurndownChartValueConfig : BurndownChartConfigProxy, IBurndownChartValueConfig
        {
            private readonly BurndownChartCategoryConfig config;

            private string name;

            private double value;

            private Action<IBarChartValue> style;

            public BurndownChartValueConfig(BurndownChartCategoryConfig config, string name, double value) :
                base(config)
            {
                this.config = config;
                this.name = name;
                this.value = value;
            }

            internal string Name => this.name;

            internal double Value => this.value;

            internal Action<IBarChartValue> Style => this.style;

            public IBurndownChartValueConfig AddValue(string name, double value)
            {
                return this.config.AddValue(name, value);
            }

            public IBurndownChartValueConfig WithStyle(Action<IBarChartValue> style)
            {
                this.style = style;
                return this;
            }
        }
    }
}
