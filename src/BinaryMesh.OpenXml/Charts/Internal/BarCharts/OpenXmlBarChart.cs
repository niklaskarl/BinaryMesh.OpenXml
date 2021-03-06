using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Drawing.Charts;

using BinaryMesh.OpenXml.Helpers;
using BinaryMesh.OpenXml.Spreadsheets;
using BinaryMesh.OpenXml.Spreadsheets.Helpers;

namespace BinaryMesh.OpenXml.Charts.Internal
{
    internal sealed class OpenXmlBarChart : IOpenXmlChart, IBarChart, IChart
    {
        private readonly OpenXmlChartSpace chartSpace;

        private readonly BarChart barChart;

        public OpenXmlBarChart(OpenXmlChartSpace chartSpace, BarChart barChart)
        {
            this.chartSpace = chartSpace;
            this.barChart = barChart;
        }

        public uint SeriesCount => (uint)this.Series.Count;

        public IReadOnlyList<IBarChartSeries> Series => new EnumerableList<BarChartSeries, IBarChartSeries>(
            this.barChart.Elements<BarChartSeries>(),
            barChartSeries => new OpenXmlBarChartSeries(barChartSeries)
        );

        public IBarChart SetDirection(BarChartDirection direction)
        {
            BarDirection barDirection = this.barChart.GetFirstChild<BarDirection>() ?? this.barChart.AppendChild(new BarDirection());
            barDirection.Val = (BarDirectionValues)direction;

            return this;
        }

        public IBarChart SetGrouping(BarChartGrouping grouping)
        {
            BarGrouping barGrouping = this.barChart.GetFirstChild<BarGrouping>() ?? this.barChart.AppendChild(new BarGrouping());
            barGrouping.Val = (BarGroupingValues)grouping;

            return this;
        }

        public IBarChart SetGapWidth(double ratio)
        {
            GapWidth gapWidth = this.barChart.GetFirstChild<GapWidth>() ?? this.barChart.AppendChild(new GapWidth());
            gapWidth.Val = (ushort)(ratio * 100);

            return this;
        }

        public IBarChart SetOverlap(double ratio)
        {
            Overlap overlap = this.barChart.GetFirstChild<Overlap>() ?? this.barChart.AppendChild(new Overlap());
            overlap.Val = (sbyte)(ratio * 100);

            return this;
        }

        public IBarChart InitializeFromRange(IRange labelRange, IRange categoryRange)
        {
            uint orderStart = (uint)this.chartSpace.Charts.TakeWhile(c => c != this).Sum(c => c.SeriesCount);

            this.barChart.RemoveAllChildren<BarChartSeries>();

            IWorksheet worksheet = labelRange.Worksheet;

            if (labelRange.Width == 1 && labelRange.Height > 0)
            {
                if (categoryRange.Height == 1 && categoryRange.Width > 0)
                {
                    for (uint labelIndex = 0; labelIndex < labelRange.Height; ++labelIndex)
                    {
                        BarChartSeries series = this.barChart.AppendChild(new BarChartSeries()  { Index = new Index() { Val = labelIndex }, Order = new Order() { Val = orderStart + labelIndex } });
                        series.AppendChild(
                            new SeriesText().AppendChildFluent(
                                new StringReference()
                                {
                                    Formula = new Formula(labelRange[0, labelIndex].Reference),
                                    StringCache = new StringCache()
                                        .AppendChildFluent(new PointCount() { Val = 1 })
                                        .AppendChildFluent(new StringPoint() { Index = 0, NumericValue = new NumericValue() { Text = labelRange[0, labelIndex].InnerValue } })
                                }
                            )
                        );

                        series.AppendChild(
                            new CategoryAxisData().AppendChildFluent(
                                new StringReference()
                                {
                                    Formula = new Formula(categoryRange.Formula),
                                    StringCache = new StringCache()
                                        .AppendChildFluent(new PointCount() { Val = (uint)categoryRange.Width.Value })
                                        .AppendFluent(Enumerable.Range(0, categoryRange.Width.Value).Select(categoryIndex => new StringPoint() { Index = (uint)categoryIndex, NumericValue = new NumericValue() { Text = categoryRange[(uint)categoryIndex, 0].InnerValue } }))
                                }
                            )
                        );

                        string valuesFormula = ReferenceEncoder.EncodeRangeReference(
                            worksheet.Name,
                            categoryRange.StartColumn, false,
                            labelRange.StartRow + labelIndex, false,
                            categoryRange.EndColumn, false,
                            labelRange.StartRow + labelIndex, false
                        );

                        series.AppendChild(
                            new Values().AppendChildFluent(
                                new NumberReference()
                                {
                                    Formula = new Formula(valuesFormula),
                                    NumberingCache = new NumberingCache()
                                        .AppendChildFluent(new PointCount() { Val = (uint)categoryRange.Width.Value })
                                        .AppendFluent(Enumerable.Range(0, categoryRange.Width.Value).Select(categoryIndex => new NumericPoint() { Index = (uint)categoryIndex, NumericValue = new NumericValue() { Text = worksheet.Cells[categoryRange.StartColumn.Value + (uint)categoryIndex, labelRange.StartRow.Value + labelIndex].InnerValue } }))
                                }
                            )
                        );
                    }
                }
                else
                {
                    throw new ArgumentException();
                }
            }
            else if (labelRange.Height == 1 && labelRange.Width > 0)
            {
                if (categoryRange.Width == 1 && categoryRange.Height > 0)
                {
                    for (uint labelIndex = 0; labelIndex < labelRange.Width; ++labelIndex)
                    {
                        BarChartSeries series = this.barChart.AppendChild(new BarChartSeries()  { Index = new Index() { Val = labelIndex }, Order = new Order() { Val = orderStart + labelIndex } });
                        series.AppendChild(
                            new SeriesText().AppendChildFluent(
                                new StringReference()
                                {
                                    Formula = new Formula(labelRange[labelIndex, 0].Reference),
                                    StringCache = new StringCache()
                                        .AppendChildFluent(new PointCount() { Val = 1 })
                                        .AppendChildFluent(new StringPoint() { Index = 0, NumericValue = new NumericValue() { Text = labelRange[labelIndex, 0].InnerValue } })
                                }
                            )
                        );

                        series.AppendChild(
                            new CategoryAxisData().AppendChildFluent(
                                new StringReference()
                                {
                                    Formula = new Formula(categoryRange.Formula),
                                    StringCache = new StringCache()
                                        .AppendChildFluent(new PointCount() { Val = (uint)categoryRange.Height.Value })
                                        .AppendFluent(Enumerable.Range(0, categoryRange.Height.Value).Select(categoryIndex => new StringPoint() { Index = (uint)categoryIndex, NumericValue = new NumericValue() { Text = categoryRange[0, (uint)categoryIndex].InnerValue } }))
                                }
                            )
                        );

                        string valuesFormula = ReferenceEncoder.EncodeRangeReference(
                            worksheet.Name,
                            labelRange.StartColumn + labelIndex, false,
                            categoryRange.StartRow, false,
                            labelRange.StartColumn + labelIndex, false,
                            categoryRange.EndRow, false
                        );

                        series.AppendChild(
                            new Values().AppendChildFluent(
                                new NumberReference()
                                {
                                    Formula = new Formula(valuesFormula),
                                    NumberingCache = new NumberingCache()
                                        .AppendChildFluent(new PointCount() { Val = (uint)categoryRange.Height.Value })
                                        .AppendFluent(Enumerable.Range(0, categoryRange.Height.Value).Select(categoryIndex => new NumericPoint() { Index = (uint)categoryIndex, NumericValue = new NumericValue() { Text = worksheet.Cells[labelRange.StartColumn.Value + labelIndex, categoryRange.StartRow.Value + (uint)categoryIndex].InnerValue } }))
                                }
                            )
                        );
                    }
                }
                else
                {
                    throw new ArgumentException();
                }
            }
            else
            {
                throw new ArgumentException();
            }

            return this;
        }
    }
}
