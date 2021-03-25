using System;
using DocumentFormat.OpenXml.Drawing.Charts;

using BinaryMesh.OpenXml.Presentations.Helpers;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal sealed class OpenXmlBarChartSeries : IBarChartSeries
    {
        private readonly BarChartSeries barChartSeries;

        public OpenXmlBarChartSeries(BarChartSeries barChartSeries)
        {
            this.barChartSeries = barChartSeries;
        }

        public IBarChartSeries SetFill(OpenXmlColor color)
        {
            ChartShapeProperties chartShapeProperties = this.barChartSeries.GetFirstChild<ChartShapeProperties>() ?? this.barChartSeries.AppendChild(new ChartShapeProperties());
            OpenXmlShapeStyler.SetFill(chartShapeProperties, color);
            return this;
        }

        public IBarChartSeries SetStroke(OpenXmlColor color)
        {
            ChartShapeProperties chartShapeProperties = this.barChartSeries.GetFirstChild<ChartShapeProperties>() ?? this.barChartSeries.AppendChild(new ChartShapeProperties());
            OpenXmlShapeStyler.SetStroke(chartShapeProperties, color);
            return this;
        }
    }
}
