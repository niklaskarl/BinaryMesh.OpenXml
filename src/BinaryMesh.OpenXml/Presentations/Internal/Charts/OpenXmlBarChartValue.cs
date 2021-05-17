using System;
using System.Linq;
using DocumentFormat.OpenXml.Drawing.Charts;

using BinaryMesh.OpenXml.Presentations.Helpers;
using BinaryMesh.OpenXml.Presentations.Internal.Mixins;
using DocumentFormat.OpenXml;

namespace BinaryMesh.OpenXml.Presentations.Internal
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
