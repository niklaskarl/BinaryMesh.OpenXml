using System;
using DocumentFormat.OpenXml.Drawing.Charts;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal sealed class OpenXmlDoughnutChart : IPieChart, IChart
    {
        private readonly PieChart pieChart;

        public OpenXmlDoughnutChart(PieChart pieChart)
        {
            this.pieChart = pieChart;
        }
        public IChartSeries Series => new OpenXmlChartSeries(this.pieChart.GetFirstChild<BarChartSeries>());
    }
}
