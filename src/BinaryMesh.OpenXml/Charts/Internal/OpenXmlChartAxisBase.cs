using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;

using BinaryMesh.OpenXml.Styles;
using BinaryMesh.OpenXml.Styles.Internal;
using BinaryMesh.OpenXml.Styles.Internal.Mixins;

namespace BinaryMesh.OpenXml.Charts.Internal
{
    internal class OpenXmlChartAxisBase : IOpenXmlShapeElement, IOpenXmlTextElement, IChartAxis
    {
        protected readonly OpenXmlElement axis;

        public OpenXmlChartAxisBase(OpenXmlElement axis)
        {
            this.axis = axis;
        }

        public uint Id => axis.GetFirstChild<AxisId>().Val;

        public IChartAxisMajorGridlines<IChartAxis> MajorGridlines => new OpenXmlChartAxisMajorGridlines<IChartAxis>(this.axis, this);

        public IVisualStyle<IChartAxis> Style => new OpenXmlVisualStyle<OpenXmlChartAxisBase, IChartAxis>(this, this);

        public ITextStyle<IChartAxis> Text => new OpenXmlTextStyle<OpenXmlChartAxisBase, IChartAxis>(this, this);

        public IChartAxis SetVisibility(bool value)
        {
            Delete delete = this.axis.GetFirstChild<Delete>() ?? this.axis.AppendChild(new Delete());
            delete.Val = !value;

            return this;
        }

        public OpenXmlElement GetOrCreateShapeProperties()
        {
            return this.axis.GetFirstChild<ShapeProperties>() ?? this.axis.AppendChild(new ShapeProperties());
        }

        public OpenXmlElement GetShapeProperties()
        {
            return this.axis.GetFirstChild<ShapeProperties>();
        }

        public OpenXmlElement GetOrCreateTextBody()
        {
            return this.axis.GetFirstChild<TextProperties>() ?? this.axis.AppendChild(new TextProperties());
        }

        public OpenXmlElement GetTextBody()
        {
            return this.axis.GetFirstChild<TextProperties>();
        }
    }
}
