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

            Index index = dataLabel.GetFirstChild<Index>();

            ChartShapeProperties chartShapeProperties = dataLabel.GetFirstChild<ChartShapeProperties>();
            if (chartShapeProperties == null)
            {
                chartShapeProperties = new ChartShapeProperties();
                dataLabel.InsertAfter(chartShapeProperties, index);
            }

            return chartShapeProperties;
        }

        public OpenXmlElement GetTextBody()
        {
            OpenXmlElement dataLabel = this.element.GetDataLabelAdjust();
            return dataLabel?.GetFirstChild<TextProperties>();
        }

        public OpenXmlElement GetOrCreateTextBody()
        {
            OpenXmlElement dataLabel = this.element.GetOrCreateDataLabelAdjust();

            Index index = dataLabel.GetFirstChild<Index>();
            ChartShapeProperties chartShapeProperties = dataLabel.GetFirstChild<ChartShapeProperties>();

            TextProperties textProperties = dataLabel.GetFirstChild<TextProperties>();
            if (textProperties == null)
            {
                textProperties = new TextProperties()
                {
                    BodyProperties = new Drawing.BodyProperties()
                    {
                        Wrap = Drawing.TextWrappingValues.Square,
                        Anchor = Drawing.TextAnchoringTypeValues.Center
                    }
                    .AppendChildFluent(new Drawing.ShapeAutoFit())
                }
                    .AppendChildFluent(new Drawing.ListStyle())
                    .AppendChildFluent(new Drawing.Paragraph());

                if (chartShapeProperties != null)
                {
                    dataLabel.InsertAfter(textProperties, chartShapeProperties);
                }
                else
                {
                    dataLabel.InsertAfter(textProperties, index);
                }
            }

            return textProperties;
        }

        public TFluent SetDelete(bool value)
        {
            OpenXmlElement dataLabelAdjust = this.element.GetOrCreateDataLabelAdjust();

            if (value)
            {
                Index index = dataLabelAdjust.GetFirstChild<Index>();
                dataLabelAdjust.RemoveAllChildren();
                dataLabelAdjust.AppendChild(index);
                dataLabelAdjust.AppendChild(new Delete() { Val = true });
            }
            else
            {
                dataLabelAdjust.RemoveAllChildren<Delete>();
            }

            return this.result;
        }
        
        public TFluent SetShowValue(bool show)
        {
            OpenXmlElement dataLabel = this.element.GetOrCreateDataLabelAdjust();
            ShowValue showValue = dataLabel.GetFirstChild<ShowValue>() ?? dataLabel.AppendChild(new ShowValue());
            showValue.Val = show;

            return this.result;
        }

        public TFluent SetShowPercent(bool show)
        {
            OpenXmlElement dataLabel = this.element.GetOrCreateDataLabelAdjust();
            ShowPercent showPercent = dataLabel.GetFirstChild<ShowPercent>() ?? dataLabel.AppendChild(new ShowPercent());
            showPercent.Val = show;

            return this.result;
        }

        public TFluent SetShowCategoryName(bool show)
        {
            OpenXmlElement dataLabel = this.element.GetOrCreateDataLabelAdjust();
            ShowCategoryName showCategoryName = dataLabel.GetFirstChild<ShowCategoryName>() ?? dataLabel.AppendChild(new ShowCategoryName());
            showCategoryName.Val = show;

            return this.result;
        }

        public TFluent SetShowLegendKey(bool show)
        {
            OpenXmlElement dataLabel = this.element.GetOrCreateDataLabelAdjust();
            ShowLegendKey showLegendKey = dataLabel.GetFirstChild<ShowLegendKey>() ?? dataLabel.AppendChild(new ShowLegendKey());
            showLegendKey.Val = show;

            return this.result;
        }

        public TFluent SetShowSeriesName(bool show)
        {
            OpenXmlElement dataLabel = this.element.GetOrCreateDataLabelAdjust();
            ShowSeriesName showSeriesName = dataLabel.GetFirstChild<ShowSeriesName>() ?? dataLabel.AppendChild(new ShowSeriesName());
            showSeriesName.Val = show;

            return this.result;
        }

        public TFluent SetShowBubbleSize(bool show)
        {
            OpenXmlElement dataLabel = this.element.GetOrCreateDataLabelAdjust();
            ShowBubbleSize showBubbleSize = dataLabel.GetFirstChild<ShowBubbleSize>() ?? dataLabel.AppendChild(new ShowBubbleSize());
            showBubbleSize.Val = show;

            return this.result;
        }

        public TFluent SetShowLeaderLines(bool show)
        {
            OpenXmlElement dataLabel = this.element.GetOrCreateDataLabelAdjust();
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
