using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;

using BinaryMesh.OpenXml.Charts.Internal.Mixins;
using BinaryMesh.OpenXml.Helpers;
using BinaryMesh.OpenXml.Styles;
using BinaryMesh.OpenXml.Styles.Internal;
using BinaryMesh.OpenXml.Styles.Internal.Mixins;

namespace BinaryMesh.OpenXml.Charts.Internal
{
    internal abstract class OpenXmlChartSeries<TSeriesFluent, TValueFluent> : IChartSeries<TSeriesFluent, TValueFluent>, IOpenXmlDataLabelElement, IOpenXmlShapeElement
        where TValueFluent : IChartValue<TValueFluent>
    {
        protected readonly OpenXmlElement series;

        public OpenXmlChartSeries(OpenXmlElement series)
        {
            this.series = series;
        }

        protected abstract TSeriesFluent Result { get; }

        internal OpenXmlElement Element => this.series;

        public IReadOnlyList<TValueFluent> Values =>
            new EnumerableList<NumericPoint, TValueFluent>(
                this.series.GetFirstChild<Values>()?.NumberReference?.NumberingCache?.Elements<NumericPoint>() ?? Enumerable.Empty<NumericPoint>(),
                p => this.ConstructValue(p.Index)
            );

        public IVisualStyle<TSeriesFluent> Style => new OpenXmlVisualStyle<OpenXmlChartSeries<TSeriesFluent, TValueFluent>, TSeriesFluent>(this, this.Result);

        public IDataLabel<TSeriesFluent> DataLabel => new OpenXmlDataLabel<OpenXmlChartSeries<TSeriesFluent, TValueFluent>, TSeriesFluent>(this, this.Result);

        protected abstract TValueFluent ConstructValue(uint index);

        public OpenXmlElement GetShapeProperties()
        {
            return this.series.GetFirstChild<ChartShapeProperties>();
        }

        public OpenXmlElement GetOrCreateShapeProperties()
        {
            return this.series.GetFirstChild<ChartShapeProperties>() ?? this.series.AppendChild(new ChartShapeProperties());
        }

        public OpenXmlElement GetDataLabel()
        {
            return this.series.GetFirstChild<DataLabels>();
        }

        public OpenXmlElement GetOrCreateDataLabel()
        {
            return this.series.GetFirstChild<DataLabels>() ?? series.AppendChild(
                new DataLabels()
                    .AppendChildFluent(new ShowLegendKey() { Val = false })
                    .AppendChildFluent(new ShowValue() { Val = false })
                    .AppendChildFluent(new ShowCategoryName() { Val = false })
                    .AppendChildFluent(new ShowSeriesName() { Val = false })
                    .AppendChildFluent(new ShowPercent() { Val = false })
                    .AppendChildFluent(new ShowBubbleSize() { Val = true })
                    .AppendChildFluent(new ShowLeaderLines() { Val = true })
            );
        }
    }
}
