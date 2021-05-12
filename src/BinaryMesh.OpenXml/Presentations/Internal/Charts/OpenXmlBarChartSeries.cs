using System;
using DocumentFormat.OpenXml.Drawing.Charts;

using BinaryMesh.OpenXml.Presentations.Helpers;
using BinaryMesh.OpenXml.Presentations.Internal.Mixins;
using DocumentFormat.OpenXml;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal sealed class OpenXmlBarChartSeries : IOpenXmlShapeElement, IBarChartSeries, IOpenXmlDataLabelElement
    {
        private readonly BarChartSeries barChartSeries;

        public OpenXmlBarChartSeries(BarChartSeries barChartSeries)
        {
            this.barChartSeries = barChartSeries;
        }

        public IVisualStyle<IBarChartSeries> Style => new OpenXmlVisualStyle<OpenXmlBarChartSeries, IBarChartSeries>(this, this);

        public IDataLabel<IBarChartSeries> DataLabel => new OpenXmlDataLabel<OpenXmlBarChartSeries, IBarChartSeries>(this, this);

        public OpenXmlElement GetShapeProperties()
        {
            return this.barChartSeries.GetFirstChild<ChartShapeProperties>();
        }

        public OpenXmlElement GetOrCreateShapeProperties()
        {
            return this.barChartSeries.GetFirstChild<ChartShapeProperties>() ?? this.barChartSeries.AppendChild(new ChartShapeProperties());
        }

        public OpenXmlElement GetDataLabel()
        {
            return this.barChartSeries.GetFirstChild<DataLabels>();
        }

        public OpenXmlElement GetOrCreateDataLabel()
        {
            return this.barChartSeries.GetFirstChild<DataLabels>() ?? this.barChartSeries.AppendChild(new DataLabels());
        }
    }
}
