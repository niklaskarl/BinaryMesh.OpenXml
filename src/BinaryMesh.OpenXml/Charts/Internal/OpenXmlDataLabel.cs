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
    internal sealed class OpenXmlDataLabel<TElement, TFluent> : IDataLabel<TFluent>, IOpenXmlShapeElement, IOpenXmlTextElement
        where TElement : IOpenXmlDataLabelElement
    {
        private readonly TElement element;

        private readonly TFluent result;

        public OpenXmlDataLabel(TElement element, TFluent result)
        {
            this.element = element;
            this.result = result;
        }

        public IVisualStyle<TFluent> Style => new OpenXmlVisualStyle<OpenXmlDataLabel<TElement, TFluent>, TFluent>(this, this.result);

        public ITextStyle<TFluent> Text => new OpenXmlTextStyle<OpenXmlDataLabel<TElement, TFluent>, TFluent>(this, this.result);

        public OpenXmlElement GetShapeProperties()
        {
            OpenXmlElement dataLabel = this.element.GetDataLabel();
            return dataLabel?.GetFirstChild<ChartShapeProperties>();
        }

        public OpenXmlElement GetOrCreateShapeProperties()
        {
            OpenXmlElement dataLabel = this.element.GetOrCreateDataLabel();
            return dataLabel.GetFirstChild<ChartShapeProperties>() ?? dataLabel.PrependChild(new ChartShapeProperties());
        }

        public OpenXmlElement GetTextBody()
        {
            OpenXmlElement dataLabel = this.element.GetDataLabel();
            return dataLabel?.GetFirstChild<TextProperties>();
        }

        public OpenXmlElement GetOrCreateTextBody()
        {
            OpenXmlElement dataLabel = this.element.GetOrCreateDataLabel();
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

        public TFluent SetShowValue(bool show)
        {
            OpenXmlElement dataLabel = this.element.GetOrCreateDataLabel();
            ShowValue showValue = dataLabel.GetFirstChild<ShowValue>() ?? dataLabel.AppendChild(new ShowValue());
            showValue.Val = show;

            return this.result;
        }

        public TFluent SetShowPercent(bool show)
        {
            OpenXmlElement dataLabel = this.element.GetOrCreateDataLabel();
            ShowPercent showPercent = dataLabel.GetFirstChild<ShowPercent>() ?? dataLabel.AppendChild(new ShowPercent());
            showPercent.Val = show;

            return this.result;
        }

        public TFluent SetShowCategoryName(bool show)
        {
            OpenXmlElement dataLabel = this.element.GetOrCreateDataLabel();
            ShowCategoryName showCategoryName = dataLabel.GetFirstChild<ShowCategoryName>() ?? dataLabel.AppendChild(new ShowCategoryName());
            showCategoryName.Val = show;

            return this.result;
        }

        public TFluent SetShowLegendKey(bool show)
        {
            OpenXmlElement dataLabel = this.element.GetOrCreateDataLabel();
            ShowLegendKey showLegendKey = dataLabel.GetFirstChild<ShowLegendKey>() ?? dataLabel.AppendChild(new ShowLegendKey());
            showLegendKey.Val = show;

            return this.result;
        }

        public TFluent SetShowSeriesName(bool show)
        {
            OpenXmlElement dataLabel = this.element.GetOrCreateDataLabel();
            ShowSeriesName showSeriesName = dataLabel.GetFirstChild<ShowSeriesName>() ?? dataLabel.AppendChild(new ShowSeriesName());
            showSeriesName.Val = show;

            return this.result;
        }

        public TFluent SetShowBubbleSize(bool show)
        {
            OpenXmlElement dataLabel = this.element.GetOrCreateDataLabel();
            ShowBubbleSize showBubbleSize = dataLabel.GetFirstChild<ShowBubbleSize>() ?? dataLabel.AppendChild(new ShowBubbleSize());
            showBubbleSize.Val = show;

            return this.result;
        }

        public TFluent SetShowLeaderLines(bool show)
        {
            OpenXmlElement dataLabel = this.element.GetOrCreateDataLabel();
            ShowLeaderLines showLeaderLines = dataLabel.GetFirstChild<ShowLeaderLines>() ?? dataLabel.AppendChild(new ShowLeaderLines());
            showLeaderLines.Val = show;

            return this.result;
        }

        public TFluent Clear()
        {
            throw new NotImplementedException();
        }
    }
}
