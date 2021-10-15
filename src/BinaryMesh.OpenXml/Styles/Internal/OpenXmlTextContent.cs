using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;

using BinaryMesh.OpenXml.Styles.Internal.Mixins;
using SixLabors.Fonts;
using BinaryMesh.OpenXml.Internal;

namespace BinaryMesh.OpenXml.Styles.Internal
{
    internal class OpenXmlTextContent<TElement, TFluent> : OpenXmlTextStyle<TElement, TFluent>, ITextContent<TFluent>, ITextStyle<TFluent>
        where TElement : IOpenXmlTextElement
    {
        public OpenXmlTextContent(TElement element, TFluent result)
            : base(element, result)
        {
        }

        public TFluent SetText(string text)
        {
            OpenXmlElement textBody = this.element.GetOrCreateTextBody();
            Paragraph paragraph = textBody.GetFirstChild<Paragraph>() ?? new Paragraph();
            Run run = paragraph.GetFirstChild<Run>() ?? new Run() { RunProperties = new RunProperties() };
            run.Text = new Text() { Text = text };

            paragraph.RemoveAllChildren<Run>();
            paragraph.AppendChild(run);

            textBody.RemoveAllChildren<Paragraph>();
            textBody.AppendChild(paragraph);

            return this.result;
        }

        public string GetText()
        {
            return this.element.GetOrCreateTextBody()?.GetFirstChild<Paragraph>()?.GetFirstChild<Run>()?.Text?.Text;
        }

        public OpenXmlSize MeasureText(IOpenXmlTextStyle defaultTextStyle, IOpenXmlTheme theme, OpenXmlUnit? width)
        {
            // TODO: subtract border
            // TODO: measure each pararaph individually

            int level = 1;
            IOpenXmlParagraphTextStyle defaultParagraphTextStyle = defaultTextStyle.GetParagraphTextStyle(level);

            string text = this.GetText();
            string typeface = theme.ResolveFontTypeface(this.GetFont() ?? defaultParagraphTextStyle.LatinTypeface);
            double fontSize = this.GetFontSize() ?? defaultParagraphTextStyle.Size ?? 9.0;
            double kerning = this.GetKerning() ?? defaultParagraphTextStyle.Kerning ?? 0.0;
            OpenXmlMargin margin =  this.GetTextMargin();

            Font font = SystemFonts.Find(typeface).CreateFont((float)fontSize);

            RendererOptions options = new RendererOptions(font, 72)
            {
                WrappingWidth = width.HasValue ? (float)(width.Value - margin.Left - margin.Right).AsPoints() : -1.0f,
                ApplyKerning = kerning != 0.0,
                LineSpacing = 1 / 1.2f
            };

            FontRectangle rect = TextMeasurer.Measure(text, options);
            OpenXmlSize result = new OpenXmlSize(
                width.HasValue ? width.Value : (OpenXmlUnit.Points(rect.Width) + margin.Left + margin.Right),
                OpenXmlUnit.Points(rect.Height * 1.2) + margin.Top + margin.Bottom
            );

            return result;
        }
    }
}
