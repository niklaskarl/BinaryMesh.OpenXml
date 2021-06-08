using System;
using DocumentFormat.OpenXml;

namespace BinaryMesh.OpenXml.Charts.Internal
{
    internal sealed class OpenXmlBarChartSeries : OpenXmlChartSeries<IBarChartSeries, IBarChartValue>, IBarChartSeries
    {
        public OpenXmlBarChartSeries(OpenXmlElement series) :
            base(series)
        {
        }

        protected override IBarChartSeries Result => this;

        protected override IChartValue<IBarChartValue> ConstructValue(uint index)
        {
            return new OpenXmlBarChartValue(this, index);
        }
    }
}
