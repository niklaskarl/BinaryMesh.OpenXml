using System;
using DocumentFormat.OpenXml;

namespace BinaryMesh.OpenXml.Charts.Internal
{
    internal sealed class OpenXmlLineChartSeries : OpenXmlChartSeries<ILineChartSeries, ILineChartValue>, ILineChartSeries
    {
        public OpenXmlLineChartSeries(OpenXmlElement series) :
            base(series)
        {
        }

        protected override ILineChartSeries Result => this;

        protected override ILineChartValue ConstructValue(uint index)
        {
            return new OpenXmlLineChartValue(this, index);
        }
    }
}
