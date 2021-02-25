using System;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;

using BinaryMesh.OpenXml.Spreadsheets;
using BinaryMesh.OpenXml.Spreadsheets.Internal;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal sealed class OpenXmlChartSpace : IChartSpace
    {
        private readonly ChartPart chartPart;

        public OpenXmlChartSpace(ChartPart chartPart)
        {
            this.chartPart = chartPart;
        }

        public ISpreadsheetDocument OpenSpreadsheetDocument()
        {
            return new OpenXmlSpreadsheetDocument(this.chartPart.EmbeddedPackagePart.GetStream());
        }

        public void AddPieChart()
        {
            this.chartPart.ChartSpace.GetFirstChild<Chart>().PlotArea
                .AppendChildFluent(
                    new PieChart()
                    {
                        VaryColors = new VaryColors() { Val = true }
                    }
                        .AppendChildFluent(
                            new DataLabels()
                            {
                                
                            }
                        )
                );
        }

        public IPieChart InsertPieChart()
        {
            throw new NotImplementedException();
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
