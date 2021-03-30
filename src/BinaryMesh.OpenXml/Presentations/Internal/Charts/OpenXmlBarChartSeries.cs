using System;
using DocumentFormat.OpenXml.Drawing.Charts;

using BinaryMesh.OpenXml.Presentations.Helpers;
using BinaryMesh.OpenXml.Presentations.Internal.Mixins;
using DocumentFormat.OpenXml;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal sealed class OpenXmlBarChartSeries : IOpenXmlShapeElement, IBarChartSeries
    {
        private readonly BarChartSeries barChartSeries;

        public OpenXmlBarChartSeries(BarChartSeries barChartSeries)
        {
            this.barChartSeries = barChartSeries;
        }

        public IVisualStyle<IBarChartSeries> Style => new OpenXmlVisualStyle<OpenXmlBarChartSeries, IBarChartSeries>(this);

        public OpenXmlElement GetShapeProperties()
        {
            return this.barChartSeries.GetFirstChild<ChartShapeProperties>();
        }

        public OpenXmlElement GetOrCreateShapeProperties()
        {
            return this.barChartSeries.GetFirstChild<ChartShapeProperties>() ?? this.barChartSeries.AppendChild(new ChartShapeProperties());
        }
    }
}
