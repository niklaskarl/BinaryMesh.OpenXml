using System;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;

using BinaryMesh.OpenXml.Charts;
using BinaryMesh.OpenXml.Charts.Internal;
using BinaryMesh.OpenXml.Styles;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    using Charts = DocumentFormat.OpenXml.Drawing.Charts;

    internal sealed class OpenXmlChartVisual : OpenXmlGraphicFrameVisual<IChartVisual>, IVisualTransform<IChartVisual>, IChartVisual
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

        protected override IChartVisual Self => this;

        public IVisualTransform<IChartVisual> Transform => this;
    }
}
