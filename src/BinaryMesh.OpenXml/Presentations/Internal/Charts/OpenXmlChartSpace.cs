using System;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;

using BinaryMesh.OpenXml.Spreadsheets;
using BinaryMesh.OpenXml.Spreadsheets.Internal;

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

        public IPieChart InsertPieChart()
        {
            ChartSpace chartSpace = this.chartPart.ChartSpace;
            Chart chart = chartSpace.GetFirstChild<Chart>() ?? chartSpace.AppendChild(new Chart());
            PlotArea plotArea = chart.PlotArea ?? (chart.PlotArea = new PlotArea());

            return new OpenXmlPieChart(
                plotArea.AppendChild(
                    new PieChart()
                        .AppendChildFluent(new BarChartSeries() { Index = new Index() { Val = 0 } })
                )
            );
        }

        public IBarChart InsertBarChart()
        {
            throw new NotImplementedException();
        }

        public IColumnChart InsertColumnChart()
        {
            throw new NotImplementedException();
        }
    }
}
