using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using Drawing = DocumentFormat.OpenXml.Drawing;

using BinaryMesh.OpenXml.Charts.Internal.Mixins;
using BinaryMesh.OpenXml.Styles;
using BinaryMesh.OpenXml.Styles.Internal;
using BinaryMesh.OpenXml.Styles.Internal.Mixins;

namespace BinaryMesh.OpenXml.Charts.Internal
{
    internal sealed class OpenXmlDataLabelAdjust<TElement, TFluent> : IValueDataLabel<TFluent>, IOpenXmlShapeElement, IOpenXmlTextElement
        where TElement : IOpenXmlDataLabelAdjustElement
    {
        private readonly TElement element;

        private readonly TFluent result;

        public OpenXmlDataLabelAdjust(TElement element, TFluent result)
        {
            this.element = element;
            this.result = result;
        }

        public IVisualStyle<TFluent> Style => new OpenXmlVisualStyle<OpenXmlDataLabelAdjust<TElement, TFluent>, TFluent>(this, this.result);

        public ITextStyle<TFluent> Text => new OpenXmlTextStyle<OpenXmlDataLabelAdjust<TElement, TFluent>, TFluent>(this, this.result);

        public OpenXmlElement GetShapeProperties()
        {
            OpenXmlElement dataLabel = this.element.GetDataLabelAdjust();
            return dataLabel?.GetFirstChild<ChartShapeProperties>();
        }

        public OpenXmlElement GetOrCreateShapeProperties()
        {
            OpenXmlElement dataLabel = this.element.GetOrCreateDataLabelAdjust();
            return dataLabel.GetFirstChild<ChartShapeProperties>() ?? dataLabel.PrependChild(new ChartShapeProperties());
        }

        public OpenXmlElement GetTextBody()
        {
            OpenXmlElement dataLabel = this.element.GetDataLabelAdjust();
            return dataLabel?.GetFirstChild<TextProperties>();
        }

        public OpenXmlElement GetOrCreateTextBody()
        {
            OpenXmlElement dataLabel = this.element.GetOrCreateDataLabelAdjust();
            TextProperties textProperties = dataLabel.GetFirstChild<TextProperties>() ?? dataLabel.PrependChild(
                new TextProperties()
                {
                    BodyProperties = new Drawing.BodyProperties()
                    {
                        Wrap = Drawing.TextWrappingValues.Square,
                        Anchor = Drawing.TextAnchoringTypeValues.Center
                    }
                    .AppendChildFluent(new Drawing.ShapeAutoFit())
                }
                .AppendChildFluent(new Drawing.ListStyle())
                .AppendChildFluent(new Drawing.Paragraph())
            );

            return textProperties;
        }

        public TFluent SetDelete(bool value)
        {
            OpenXmlElement dataLabelAdjust = this.element.GetOrCreateDataLabelAdjust();
            Delete delete = dataLabelAdjust.GetFirstChild<Delete>() ?? dataLabelAdjust.AppendChild(new Delete());
            delete.Val = value;

            return this.result;
        }

        public TFluent Clear()
        {
            throw new NotImplementedException();
        }
    }
}
