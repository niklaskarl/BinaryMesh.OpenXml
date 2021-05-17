using System;
using System.Linq;
using DocumentFormat.OpenXml.Drawing.Charts;

using BinaryMesh.OpenXml.Presentations.Helpers;
using BinaryMesh.OpenXml.Presentations.Internal.Mixins;
using DocumentFormat.OpenXml;

namespace BinaryMesh.OpenXml.Presentations.Internal
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
