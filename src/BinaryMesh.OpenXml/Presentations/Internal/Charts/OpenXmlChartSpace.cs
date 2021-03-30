using System;
using System.Linq;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using Drawing = DocumentFormat.OpenXml.Drawing;

using BinaryMesh.OpenXml.Spreadsheets;
using BinaryMesh.OpenXml.Spreadsheets.Internal;
using System.Collections.Generic;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal sealed class OpenXmlChartSpace : IOpenXmlChartSpace, IChartSpace
    {
        private readonly ChartPart chartPart;

        public OpenXmlChartSpace(ChartPart chartPart)
        {
            this.chartPart = chartPart;
        }

        public ChartPart ChartPart => this.chartPart;

        public IReadOnlyList<IChartAxis> CategoryAxes
        {
            get
            {
                IEnumerable<CategoryAxis> axes = this.chartPart.ChartSpace
                    ?.GetFirstChild<Chart>()
                    ?.GetFirstChild<PlotArea>()
                    ?.GetFirstChild<ValueAxis>()
                    ?.Elements<CategoryAxis>() ?? Enumerable.Empty<CategoryAxis>();

                return axes.Select(axis => new OpenXmlChartAxisBase(axis)).ToList();
            }
        }

        public IReadOnlyList<IChartAxis> ValueAxes
        {
            get
            {
                IEnumerable<ValueAxis> axes = this.chartPart.ChartSpace
                    ?.GetFirstChild<Chart>()
                    ?.GetFirstChild<PlotArea>()
                    ?.GetFirstChild<ValueAxis>()
                    ?.Elements<ValueAxis>() ?? Enumerable.Empty<ValueAxis>();

                return axes.Select(axis => new OpenXmlChartAxisBase(axis)).ToList();
            }
        }

        public ISpreadsheetDocument OpenSpreadsheetDocument()
        {
            if (this.chartPart.EmbeddedPackagePart == null)
            {
                string nextId = this.chartPart.GetNextRelationshipId();
                EmbeddedPackagePart embeddedPackagePart = this.chartPart.AddEmbeddedPackagePart("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
                this.chartPart.ChangeIdOfPart(embeddedPackagePart, nextId);
                this.chartPart.ChartSpace
                    .AppendChildFluent(new ExternalData()
                    {
                        Id = this.chartPart.GetIdOfPart(embeddedPackagePart),
                        AutoUpdate = new AutoUpdate() { Val = false }
                    });

                return new OpenXmlSpreadsheetDocument(embeddedPackagePart.GetStream(), true);
            }
            else
            {
                return new OpenXmlSpreadsheetDocument(this.chartPart.EmbeddedPackagePart.GetStream(), false);
            }
        }

        public CartesianAxes AppendCartesianAxes()
        {
            CategoryAxis categoryAxis = this.AppendCategoryAxis();
            ValueAxis valueAxis = this.AppendValueAxis();

            categoryAxis.CrossingAxis = new CrossingAxis() { Val = valueAxis.AxisId.Val };
            valueAxis.CrossingAxis = new CrossingAxis() { Val = categoryAxis.AxisId.Val };

            return new CartesianAxes(new OpenXmlChartAxisBase(categoryAxis), new OpenXmlChartAxisBase(valueAxis));
        }

        public IPieChart InsertPieChart()
        {
            ChartSpace chartSpace = this.chartPart.ChartSpace;
            Chart chart = chartSpace.GetFirstChild<Chart>() ?? chartSpace.AppendChild(new Chart());
            PlotArea plotArea = chart.PlotArea ?? (chart.PlotArea = new PlotArea());

            return new OpenXmlPieChart(
                plotArea.AppendChild(
                    new DoughnutChart()
                        .AppendChildFluent(new PieChartSeries() { Index = new Index() { Val = 0 } })
                )
            );
        }

        public IBarChart InsertBarChart(CartesianAxes axes)
        {
            ChartSpace chartSpace = this.chartPart.ChartSpace;
            Chart chart = chartSpace.GetFirstChild<Chart>() ?? chartSpace.AppendChild(new Chart());
            PlotArea plotArea = chart.PlotArea ?? (chart.PlotArea = new PlotArea());

            return new OpenXmlBarChart(
                plotArea.AppendChild(
                    new BarChart()
                        .AppendChildFluent(new AxisId() { Val = axes.CategoryAxis.Id })
                        .AppendChildFluent(new AxisId() { Val = axes.ValueAxis.Id })
                )
            );
        }

        private CategoryAxis AppendCategoryAxis()
        {
            ChartSpace chartSpace = this.chartPart.ChartSpace;
            Chart chart = chartSpace.GetFirstChild<Chart>() ?? chartSpace.AppendChild(new Chart());
            PlotArea plotArea = chart.GetFirstChild<PlotArea>() ?? chart.AppendChild(new PlotArea());

            uint id = plotArea.Elements<CategoryAxis>().Select(axis => axis.AxisId.Val.Value).DefaultIfEmpty(417317351u).Max() + 1;

            CategoryAxis categoryAxis = plotArea.AppendChild(
                new CategoryAxis()
                {
                    AxisId = new AxisId() { Val = id },
                    Scaling = new Scaling() { Orientation = new Orientation() { Val = OrientationValues.MinMax } },
                    Delete = new Delete() { Val = false },
                    AxisPosition = new AxisPosition() { Val = AxisPositionValues.Bottom },
                    NumberingFormat = new NumberingFormat() { FormatCode = "General", SourceLinked = true },
                    MajorTickMark = new MajorTickMark() { Val = TickMarkValues.None },
                    MinorTickMark = new MinorTickMark() { Val = TickMarkValues.None },
                    TickLabelPosition = new TickLabelPosition() { Val = TickLabelPositionValues.NextTo },
                    ChartShapeProperties = new ChartShapeProperties().AppendChildFluent(new Drawing.NoFill()), // TODO: stroke
                    TextProperties = new TextProperties()
                    {
                        BodyProperties = new Drawing.BodyProperties(),
                        ListStyle = new Drawing.ListStyle()
                    }
                        .AppendChildFluent(
                            new Drawing.Paragraph()
                            {
                                ParagraphProperties = new Drawing.ParagraphProperties().AppendChildFluent(
                                    new Drawing.DefaultRunProperties()
                                    {
                                        FontSize = 1197,
                                        Bold = false,
                                        Italic = false,
                                        Underline = Drawing.TextUnderlineValues.None,
                                        Strike = Drawing.TextStrikeValues.NoStrike,
                                        Kerning = 1200,
                                        Baseline = 0
                                    }
                                        .AppendChildFluent(
                                            new Drawing.SolidFill()
                                            {
                                                SchemeColor = new Drawing.SchemeColor() { Val = Drawing.SchemeColorValues.Text1 }
                                                    .AppendChildFluent(new Drawing.LuminanceModulation() { Val = 65000 })
                                                    .AppendChildFluent(new Drawing.LuminanceOffset() { Val = 35000 })
                                            }
                                        )
                                        .AppendChildFluent(new Drawing.LatinFont() { Typeface = "+mn-lt" })
                                        .AppendChildFluent(new Drawing.EastAsianFont() { Typeface = "+mn-ea" })
                                        .AppendChildFluent(new Drawing.ComplexScriptFont() { Typeface = "+mn-cs" })
                                )
                            }
                        )
                }
                    .AppendChildFluent(new Crosses() { Val = CrossesValues.AutoZero })
                    .AppendChildFluent(new AutoLabeled() { Val = true })
                    .AppendChildFluent(new LabelAlignment() { Val = LabelAlignmentValues.Center })
                    .AppendChildFluent(new LabelOffset() { Val = 100 })
                    .AppendChildFluent(new NoMultiLevelLabels() { Val = false })
            );

            return categoryAxis;
        }

        private ValueAxis AppendValueAxis()
        {
            ChartSpace chartSpace = this.chartPart.ChartSpace;
            Chart chart = chartSpace.GetFirstChild<Chart>() ?? chartSpace.AppendChild(new Chart());
            PlotArea plotArea = chart.GetFirstChild<PlotArea>() ?? chart.AppendChild(new PlotArea());

            uint id = plotArea.Elements<ValueAxis>().Select(axis => axis.AxisId.Val.Value).DefaultIfEmpty(417314071u).Max() + 1;

            ValueAxis valueAxis = plotArea.AppendChild(
                new ValueAxis()
                {
                    AxisId = new AxisId() { Val = id },
                    Scaling = new Scaling() { Orientation = new Orientation() { Val = OrientationValues.MinMax } },
                    Delete = new Delete() { Val = false },
                    AxisPosition = new AxisPosition() { Val = AxisPositionValues.Left },
                    MajorGridlines = new MajorGridlines()
                    {
                        ChartShapeProperties = new ChartShapeProperties()
                        {
                            // TODO
                        }
                    },
                    NumberingFormat = new NumberingFormat() { FormatCode = "General", SourceLinked = true },
                    MajorTickMark = new MajorTickMark() { Val = TickMarkValues.None },
                    MinorTickMark = new MinorTickMark() { Val = TickMarkValues.None },
                    TickLabelPosition = new TickLabelPosition() { Val = TickLabelPositionValues.NextTo },
                    ChartShapeProperties = new ChartShapeProperties()
                        .AppendChildFluent(new Drawing.NoFill())
                        .AppendChildFluent(new Drawing.Outline().AppendChildFluent(new Drawing.NoFill()))
                        .AppendChildFluent(new Drawing.EffectList()),
                    TextProperties = new TextProperties()
                    {
                        BodyProperties = new Drawing.BodyProperties(),
                        ListStyle = new Drawing.ListStyle()
                    }
                        .AppendChildFluent(
                            new Drawing.Paragraph()
                            {
                                ParagraphProperties = new Drawing.ParagraphProperties().AppendChildFluent(
                                    new Drawing.DefaultRunProperties()
                                    {
                                        FontSize = 1197,
                                        Bold = false,
                                        Italic = false,
                                        Underline = Drawing.TextUnderlineValues.None,
                                        Strike = Drawing.TextStrikeValues.NoStrike,
                                        Kerning = 1200,
                                        Baseline = 0
                                    }
                                        .AppendChildFluent(
                                            new Drawing.SolidFill()
                                            {
                                                SchemeColor = new Drawing.SchemeColor() { Val = Drawing.SchemeColorValues.Text1 }
                                                    .AppendChildFluent(new Drawing.LuminanceModulation() { Val = 65000 })
                                                    .AppendChildFluent(new Drawing.LuminanceOffset() { Val = 35000 })
                                            }
                                        )
                                        .AppendChildFluent(new Drawing.LatinFont() { Typeface = "+mn-lt" })
                                        .AppendChildFluent(new Drawing.EastAsianFont() { Typeface = "+mn-ea" })
                                        .AppendChildFluent(new Drawing.ComplexScriptFont() { Typeface = "+mn-cs" })
                                )
                            }
                        )
                }
                    .AppendChildFluent(new Crosses() { Val = CrossesValues.AutoZero })
                    .AppendChildFluent(new CrossBetween() { Val = CrossBetweenValues.Between })
            );

            return valueAxis;
        }
    }
}
