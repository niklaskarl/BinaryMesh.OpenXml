using System;

using BinaryMesh.OpenXml.Charts.Internal.Mixins;

namespace BinaryMesh.OpenXml.Charts.Internal
{
    internal sealed class OpenXmlLineChartValue : OpenXmlChartValue<ILineChartSeries, ILineChartValue>, ILineChartValue, IOpenXmlDataLabelAdjustElement
    {
        public OpenXmlLineChartValue(OpenXmlLineChartSeries series, uint valueIndex) :
            base(series, valueIndex)
        {
        }

        protected override ILineChartValue Result => this;
    }
}
