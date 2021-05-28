using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;

using BinaryMesh.OpenXml.Styles;
using BinaryMesh.OpenXml.Styles.Internal;
using BinaryMesh.OpenXml.Styles.Internal.Mixins;

namespace BinaryMesh.OpenXml.Charts.Internal
{
    internal sealed class OpenXmlChartAxisMajorGridlines<TResult> : IChartAxisMajorGridlines<TResult>, IOpenXmlShapeElement
    {
        private readonly OpenXmlElement axis;

        private readonly TResult result;

        public OpenXmlChartAxisMajorGridlines(OpenXmlElement axis, TResult result)
        {
            this.axis = axis;
            this.result = result;
        }

        public bool Exists => this.axis.GetFirstChild<MajorGridlines>() != null;

        public IStrokeStyle<TResult> Style => new OpenXmlVisualStyle<OpenXmlChartAxisMajorGridlines<TResult>, TResult>(this, this.result);

        public TResult Add()
        {
            this.GetOrCreateDefaultMajorGridlines();
            return this.result;
        }

        public TResult Remove()
        {
            this.axis.RemoveAllChildren<MajorGridlines>();
            return this.result;
        }

        public OpenXmlElement GetShapeProperties()
        {
            return this.axis.GetFirstChild<MajorGridlines>()?.GetFirstChild<ChartShapeProperties>();
        }

        public OpenXmlElement GetOrCreateShapeProperties()
        {
            OpenXmlElement gridlines = this.GetOrCreateDefaultMajorGridlines();
            return gridlines.GetFirstChild<ChartShapeProperties>() ?? gridlines.AppendChild(new ChartShapeProperties());
        }

        private OpenXmlElement GetOrCreateDefaultMajorGridlines()
        {
            return this.axis.GetFirstChild<MajorGridlines>() ?? this.axis.AppendChild(new MajorGridlines());
        }
    }
}
