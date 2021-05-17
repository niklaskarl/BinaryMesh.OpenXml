using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using Drawing = DocumentFormat.OpenXml.Drawing;

using BinaryMesh.OpenXml.Spreadsheets;
using System.Linq;
using System.Collections.Generic;
using BinaryMesh.OpenXml.Presentations.Internal.Mixins;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal abstract class OpenXmlChartValue<TSeriesFluent, TValueFluent> : IChartValue<TValueFluent>, IOpenXmlShapeElement, IOpenXmlDataLabelAdjustElement
        where TValueFluent : IChartValue<TValueFluent>
    {
        protected readonly OpenXmlChartSeries<TSeriesFluent, TValueFluent> series;

        protected readonly uint valueIndex;

        public OpenXmlChartValue(OpenXmlChartSeries<TSeriesFluent, TValueFluent> series, uint valueIndex)
        {
            this.series = series;
            this.valueIndex = valueIndex;
        }

        protected abstract TValueFluent Result { get; }

        public IVisualStyle<TValueFluent> Style => new OpenXmlVisualStyle<OpenXmlChartValue<TSeriesFluent, TValueFluent>, TValueFluent>(this, this.Result);

        public IDataLabelAdjust<TValueFluent> DataLabel => new OpenXmlDataLabelAdjust<OpenXmlChartValue<TSeriesFluent, TValueFluent>, TValueFluent>(this, this.Result);

        public OpenXmlElement GetDataLabelAdjust()
        {
            return this.series.GetDataLabel()?.Elements<DataLabel>().FirstOrDefault(dl => dl.Index?.Val == valueIndex);
        }

        public OpenXmlElement GetOrCreateDataLabelAdjust()
        {
            OpenXmlElement dataLabel = this.series.GetOrCreateDataLabel();
            DataLabel dataLabelAdjust = dataLabel.Elements<DataLabel>().Where(dl => dl.Index?.Val <= this.valueIndex).LastOrDefault();
            if (!(dataLabelAdjust?.Index?.Val?.HasValue ?? false) || dataLabelAdjust.Index.Val != this.valueIndex)
            {
                dataLabelAdjust = dataLabel.InsertAfter(new DataLabel() { Index = new Index() { Val = this.valueIndex} }, dataLabelAdjust);
            }

            return dataLabelAdjust;
        }

        public OpenXmlElement GetShapeProperties()
        {
            return this.series.Element.Elements<DataPoint>().FirstOrDefault(dl => dl.Index?.Val == valueIndex)?.ChartShapeProperties;
        }

        public OpenXmlElement GetOrCreateShapeProperties()
        {
            DataPoint dataPoint = this.series.Element.Elements<DataPoint>().Where(dp => dp.Index?.Val <= this.valueIndex).LastOrDefault();
            if (!(dataPoint?.Index?.Val?.HasValue ?? false) || dataPoint.Index.Val != this.valueIndex)
            {
                dataPoint = this.series.Element.InsertAfter(
                    new DataPoint()
                    {
                        Index = new Index() { Val = this.valueIndex },
                        Bubble3D = new Bubble3D() { Val = false }
                    },
                    dataPoint
                );
            }
            
            return dataPoint.ChartShapeProperties ?? dataPoint.AppendChild(new ChartShapeProperties());
        }
    }
}
