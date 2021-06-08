using System;

using BinaryMesh.OpenXml.Charts.Internal.Mixins;

namespace BinaryMesh.OpenXml.Charts.Internal
{
    internal sealed class OpenXmlBarChartValue : OpenXmlChartValue<IBarChartSeries, IBarChartValue>, IBarChartValue, IOpenXmlDataLabelAdjustElement
    {
        public OpenXmlBarChartValue(OpenXmlBarChartSeries series, uint valueIndex) :
            base(series, valueIndex)
        {
        }

        protected override IBarChartValue Result => this;
    }
}
