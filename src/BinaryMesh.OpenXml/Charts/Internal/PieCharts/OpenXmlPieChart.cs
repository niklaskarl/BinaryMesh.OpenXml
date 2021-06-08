using System;
using DocumentFormat.OpenXml.Drawing.Charts;

namespace BinaryMesh.OpenXml.Charts.Internal
{
    internal sealed class OpenXmlPieChart : IOpenXmlChart, IPieChart, IChart
    {
        private readonly OpenXmlChartSpace chartSpace;

        private readonly DoughnutChart doughnutChart;

        public OpenXmlPieChart(OpenXmlChartSpace chartSpace, DoughnutChart doughnutChart)
        {
            this.chartSpace = chartSpace;
            this.doughnutChart = doughnutChart;
        }

        public uint SeriesCount => 1;

        public IPieChartSeries Series => new OpenXmlPieChartSeries(this.doughnutChart.GetFirstChild<PieChartSeries>());

        public IPieChart SetFirstSliceAngle(double rad)
        {
            FirstSliceAngle firstSliceAngle = this.doughnutChart.GetFirstChild<FirstSliceAngle>() ?? this.doughnutChart.AppendChild(new FirstSliceAngle());
            firstSliceAngle.Val = (ushort)((rad * 180) / Math.PI);

            return this;
        }

        public IPieChart SetExplosion(double percent)
        {
            PieChartSeries pieChartSeries = this.doughnutChart.GetFirstChild<PieChartSeries>() ?? this.doughnutChart.AppendChild(new PieChartSeries());
            Explosion explosion = pieChartSeries.GetFirstChild<Explosion>() ?? pieChartSeries.AppendChild(new Explosion());
            explosion.Val = (ushort)(percent * 100);

            return this;
        }

        public IPieChart SetHoleSize(double percent)
        {
            HoleSize holeSize = this.doughnutChart.GetFirstChild<HoleSize>() ?? this.doughnutChart.AppendChild(new HoleSize());
            holeSize.Val = (byte)(percent * 100);

            return this;
        }
    }
}
