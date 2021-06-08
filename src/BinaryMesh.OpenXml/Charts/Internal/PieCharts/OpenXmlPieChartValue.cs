using System;

using BinaryMesh.OpenXml.Charts.Internal.Mixins;

namespace BinaryMesh.OpenXml.Charts.Internal
{
    internal sealed class OpenXmlPieChartValue : OpenXmlChartValue<IPieChartSeries, IPieChartValue>, IPieChartValue, IOpenXmlDataLabelAdjustElement
    {
        public OpenXmlPieChartValue(OpenXmlPieChartSeries series, uint valueIndex) :
            base(series, valueIndex)
        {
        }

        protected override IPieChartValue Result => this;
    }
}
