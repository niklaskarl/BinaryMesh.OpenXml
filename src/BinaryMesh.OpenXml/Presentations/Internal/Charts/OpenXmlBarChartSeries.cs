using System;
using DocumentFormat.OpenXml.Drawing.Charts;

using BinaryMesh.OpenXml.Presentations.Helpers;
using BinaryMesh.OpenXml.Presentations.Internal.Mixins;
using DocumentFormat.OpenXml;
using System.Collections.Generic;
using BinaryMesh.OpenXml.Helpers;
using System.Linq;

namespace BinaryMesh.OpenXml.Presentations.Internal
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
