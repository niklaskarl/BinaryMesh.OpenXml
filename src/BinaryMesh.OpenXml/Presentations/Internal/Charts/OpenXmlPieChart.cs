using System;
using DocumentFormat.OpenXml.Drawing.Charts;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal sealed class OpenXmlPieChart : IPieChart, IChart
    {
        private readonly PieChart pieChart;

        public OpenXmlPieChart(PieChart pieChart)
        {
            this.pieChart = pieChart;
        }
        public IChartSeries Series => new OpenXmlChartSeries(this.pieChart.GetFirstChild<BarChartSeries>());
    }
}
