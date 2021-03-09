using System;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;
using Charts = DocumentFormat.OpenXml.Drawing.Charts;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal sealed class OpenXmlChartVisual : OpenXmlGraphicFrameVisual, IChartVisual
    {
        private readonly Charts.Chart chart;

        public OpenXmlChartVisual(IOpenXmlVisualContainer container, GraphicFrame graphicFrame) :
            base(container, graphicFrame)
        {
            this.chart = this.graphicFrame
                .GetFirstChild<Drawing.Graphic>()
                .GetFirstChild<Drawing.GraphicData>()
                .GetFirstChild<Charts.Chart>();
        }

        public IChartSpace ChartSpace =>
            new OpenXmlChartSpace(this.container.Part.GetPartById(this.chart.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships").Value) as ChartPart);

        IChartVisual IChartVisual.SetExtents(long width, long height)
        {
            this.SetExtents(width, height);
            return this;
        }

        IChartVisual IChartVisual.SetOffset(long x, long y)
        {
            this.SetOffset(x, y);
            return this;
        }
    }
}
