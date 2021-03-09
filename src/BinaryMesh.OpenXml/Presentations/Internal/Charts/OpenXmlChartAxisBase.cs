using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal class OpenXmlChartAxisBase : IChartAxis
    {
        protected readonly OpenXmlElement axis;

        public OpenXmlChartAxisBase(OpenXmlElement axis)
        {
            this.axis = axis;
        }

        public uint Id => axis.GetFirstChild<AxisId>().Val;
    }
}
