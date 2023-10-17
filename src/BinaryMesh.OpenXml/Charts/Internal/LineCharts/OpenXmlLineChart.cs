using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Drawing.Charts;

using BinaryMesh.OpenXml.Helpers;
using BinaryMesh.OpenXml.Spreadsheets;
using BinaryMesh.OpenXml.Spreadsheets.Helpers;

namespace BinaryMesh.OpenXml.Charts.Internal
{
    internal sealed class OpenXmlLineChart : IOpenXmlChart, ILineChart, IChart
    {
        private readonly OpenXmlChartSpace chartSpace;

        private readonly LineChart lineChart;

        public OpenXmlLineChart(OpenXmlChartSpace chartSpace, LineChart lineChart)
        {
            this.chartSpace = chartSpace;
            this.lineChart = lineChart;
        }

        public uint SeriesCount => (uint)this.Series.Count;

        public IReadOnlyList<ILineChartSeries> Series => new EnumerableList<LineChartSeries, ILineChartSeries>(
            this.lineChart.Elements<LineChartSeries>(),
            lineChartSeries => new OpenXmlLineChartSeries(lineChartSeries)
        );

        public ILineChart SetGrouping(LineChartGrouping grouping)
        {
            Grouping lineGrouping = this.lineChart.GetFirstChild<Grouping>() ?? this.lineChart.AppendChild(new Grouping());
            lineGrouping.Val = (GroupingValues)grouping;

            return this;
        }

        public ILineChart InitializeFromRange(IRange labelRange, IRange categoryRange)
        {
            uint orderStart = (uint)this.chartSpace.Charts.TakeWhile(c => c != this).Sum(c => c.SeriesCount);

            this.lineChart.RemoveAllChildren<LineChartSeries>();

            IWorksheet worksheet = labelRange.Worksheet;

            if (labelRange.Width == 1 && labelRange.Height > 0 && categoryRange.Height == 1 && categoryRange.Width > 0)
            {
                for (uint labelIndex = 0; labelIndex < labelRange.Height; ++labelIndex)
                {
                    LineChartSeries series = this.lineChart.AppendChild(
                        new LineChartSeries()
                        {
                            Index = new Index() { Val = orderStart + labelIndex },
                            Order = new Order() { Val = orderStart + labelIndex }
                        }
                    );

                    series.AppendChild(
                        new Marker()
                        {
                            Symbol = new Symbol() { Val = MarkerStyleValues.None }
                        }
                    );

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

                    series.AppendChild(new Smooth() { Val = false });
                }
            }
            else if (labelRange.Height == 1 && labelRange.Width > 0 && categoryRange.Width == 1 && categoryRange.Height > 0)
            {
                for (uint labelIndex = 0; labelIndex < labelRange.Width; ++labelIndex)
                {
                    LineChartSeries series = this.lineChart.AppendChild(
                        new LineChartSeries()
                        {
                            Index = new Index() { Val = orderStart + labelIndex },
                            Order = new Order() { Val = orderStart + labelIndex }
                        }
                    );

                    series.AppendChild(
                        new Marker()
                        {
                            Symbol = new Symbol() { Val = MarkerStyleValues.None }
                        }
                    );

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

                    series.AppendChild(new Smooth() { Val = false });
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
