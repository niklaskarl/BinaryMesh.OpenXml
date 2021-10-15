using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;

using BinaryMesh.OpenXml.Styles.Internal.Mixins;

namespace BinaryMesh.OpenXml.Styles.Internal
{
    internal class OpenXmlTextStyle<TElement, TFluent> : ITextStyle<TFluent>
        where TElement : IOpenXmlTextElement
    {
        protected readonly TElement element;

        protected readonly TFluent result;

        public OpenXmlTextStyle(TElement element, TFluent result)
        {
            this.element = element;
            this.result = result;
        }

        public TFluent SetFontSize(double fontSize)
        {
            OpenXmlElement textBody = this.element.GetOrCreateTextBody();
            BodyProperties bodyProperties = textBody.GetFirstChild<BodyProperties>() ?? textBody.AppendChild(new BodyProperties());

            foreach (Paragraph paragraph in textBody.Elements<Paragraph>())
            {
                ParagraphProperties paragraphProperties = paragraph.GetFirstChild<ParagraphProperties>() ?? paragraph.PrependChild(new ParagraphProperties());
                DefaultRunProperties defaultRunProperties = paragraphProperties.GetFirstChild<DefaultRunProperties>() ?? paragraphProperties.AppendChild(new DefaultRunProperties());
                defaultRunProperties.FontSize = (int)(fontSize * 100);

                foreach (Run run in paragraph.Elements<Run>())
                {
                    RunProperties runProperties = run.GetFirstChild<RunProperties>() ?? run.PrependChild(new RunProperties());
                    runProperties.FontSize = (int)(fontSize * 100);
                }
            }

            return this.result;
        }

        public double? GetFontSize()
        {
            OpenXmlElement textBody = this.element.GetTextBody();
            Int32Value value = textBody.GetFirstChild<Paragraph>()?.GetFirstChild<ParagraphProperties>()?.GetFirstChild<DefaultRunProperties>()?.FontSize;
            return (value?.HasValue ?? false) ? value.Value / 100.0 : (double?)null;
        }

        public TFluent SetFont(string typeface)
        {
            OpenXmlElement textBody = this.element.GetOrCreateTextBody();
            BodyProperties bodyProperties = textBody.GetFirstChild<BodyProperties>() ?? textBody.AppendChild(new BodyProperties());

            foreach (Paragraph paragraph in textBody.Elements<Paragraph>())
            {
                ParagraphProperties paragraphProperties = paragraph.GetFirstChild<ParagraphProperties>() ?? paragraph.PrependChild(new ParagraphProperties());
                DefaultRunProperties defaultRunProperties = paragraphProperties.GetFirstChild<DefaultRunProperties>() ?? paragraphProperties.AppendChild(new DefaultRunProperties());
                defaultRunProperties.RemoveAllChildren<LatinFont>();
                defaultRunProperties.RemoveAllChildren<ComplexScriptFont>();

                defaultRunProperties.AppendChild(new LatinFont() { Typeface = typeface });
                defaultRunProperties.AppendChild(new ComplexScriptFont() { Typeface = typeface });

                foreach (Run run in paragraph.Elements<Run>())
                {
                    RunProperties runProperties = run.GetFirstChild<RunProperties>() ?? run.PrependChild(new RunProperties());
                    runProperties.RemoveAllChildren<LatinFont>();
                    runProperties.RemoveAllChildren<ComplexScriptFont>();

                    runProperties.AppendChild(new LatinFont() { Typeface = typeface });
                    runProperties.AppendChild(new ComplexScriptFont() { Typeface = typeface });
                }
            }

            return this.result;
        }

        public string GetFont()
        {
            OpenXmlElement textBody = this.element.GetTextBody();
            return textBody.GetFirstChild<Paragraph>()?.GetFirstChild<ParagraphProperties>()?.GetFirstChild<DefaultRunProperties>()?.GetFirstChild<LatinFont>()?.Typeface;
        }

        public TFluent SetFontColor(OpenXmlColor color)
        {
            OpenXmlElement textBody = this.element.GetOrCreateTextBody();
            BodyProperties bodyProperties = textBody.GetFirstChild<BodyProperties>() ?? textBody.AppendChild(new BodyProperties());

            foreach (Paragraph paragraph in textBody.Elements<Paragraph>())
            {
                ParagraphProperties paragraphProperties = paragraph.GetFirstChild<ParagraphProperties>() ?? paragraph.PrependChild(new ParagraphProperties());
                DefaultRunProperties defaultRunProperties = paragraphProperties.GetFirstChild<DefaultRunProperties>() ?? paragraphProperties.AppendChild(new DefaultRunProperties());
                defaultRunProperties.RemoveAllChildren<NoFill>();
                defaultRunProperties.RemoveAllChildren<SolidFill>();
                defaultRunProperties.RemoveAllChildren<GradientFill>();
                defaultRunProperties.RemoveAllChildren<BlipFill>();
                defaultRunProperties.RemoveAllChildren<PatternFill>();
                defaultRunProperties.RemoveAllChildren<GroupFill>();

                defaultRunProperties.AppendChild(new SolidFill().AppendChildFluent(color.CreateColorElement()));

                foreach (Run run in paragraph.Elements<Run>())
                {
                    RunProperties runProperties = run.GetFirstChild<RunProperties>() ?? run.PrependChild(new RunProperties());
                    runProperties.RemoveAllChildren<NoFill>();
                    runProperties.RemoveAllChildren<SolidFill>();
                    runProperties.RemoveAllChildren<GradientFill>();
                    runProperties.RemoveAllChildren<BlipFill>();
                    runProperties.RemoveAllChildren<PatternFill>();
                    runProperties.RemoveAllChildren<GroupFill>();

                    runProperties.AppendChild(new SolidFill().AppendChildFluent(color.CreateColorElement()));
                }
            }

            return this.result;
        }

        public double? GetKerning()
        {
            OpenXmlElement textBody = this.element.GetTextBody();
            Int32Value value = textBody.GetFirstChild<Paragraph>()?.GetFirstChild<ParagraphProperties>()?.GetFirstChild<DefaultRunProperties>()?.Kerning;
            return (value?.HasValue ?? false) ? value.Value / 100.0 : (double?)null;
        }

        public TFluent SetTextAlign(TextAlignmentTypeValues align)
        {
            OpenXmlElement textBody = this.element.GetOrCreateTextBody();

            foreach (Paragraph paragraph in textBody.Elements<Paragraph>())
            {
                ParagraphProperties paragraphProperties = paragraph.GetFirstChild<ParagraphProperties>() ?? paragraph.PrependChild(new ParagraphProperties());
                paragraphProperties.Alignment = align;
            }

            return this.result;
        }

        public TFluent SetTextAnchor(TextAnchoringTypeValues anchor)
        {
            OpenXmlElement textBody = this.element.GetOrCreateTextBody();
            BodyProperties bodyProperties = textBody.GetFirstChild<BodyProperties>() ?? textBody.AppendChild(new BodyProperties());
            bodyProperties.Anchor = anchor;

            return this.result;
        }

        public TFluent SetIsBold(bool bold)
        {
            OpenXmlElement textBody = this.element.GetOrCreateTextBody();
            BodyProperties bodyProperties = textBody.GetFirstChild<BodyProperties>() ?? textBody.AppendChild(new BodyProperties());

            foreach (Paragraph paragraph in textBody.Elements<Paragraph>())
            {
                ParagraphProperties paragraphProperties = paragraph.GetFirstChild<ParagraphProperties>() ?? paragraph.PrependChild(new ParagraphProperties());
                DefaultRunProperties defaultRunProperties = paragraphProperties.GetFirstChild<DefaultRunProperties>() ?? paragraphProperties.AppendChild(new DefaultRunProperties());
                defaultRunProperties.Bold = bold;

                foreach (Run run in paragraph.Elements<Run>())
                {
                    RunProperties runProperties = run.GetFirstChild<RunProperties>() ?? run.PrependChild(new RunProperties());
                    runProperties.Bold = bold;
                }
            }

            return this.result;
        }

        public TFluent SetIsItalic(bool italic)
        {
            OpenXmlElement textBody = this.element.GetOrCreateTextBody();
            BodyProperties bodyProperties = textBody.GetFirstChild<BodyProperties>() ?? textBody.AppendChild(new BodyProperties());

            foreach (Paragraph paragraph in textBody.Elements<Paragraph>())
            {
                ParagraphProperties paragraphProperties = paragraph.GetFirstChild<ParagraphProperties>() ?? paragraph.PrependChild(new ParagraphProperties());
                DefaultRunProperties defaultRunProperties = paragraphProperties.GetFirstChild<DefaultRunProperties>() ?? paragraphProperties.AppendChild(new DefaultRunProperties());
                defaultRunProperties.Italic = italic;

                foreach (Run run in paragraph.Elements<Run>())
                {
                    RunProperties runProperties = run.GetFirstChild<RunProperties>() ?? run.PrependChild(new RunProperties());
                    runProperties.Italic = italic;
                }
            }

            return this.result;
        }

        public TFluent SetTextMargin(long left, long top, long right, long bottom)
        {
            OpenXmlElement textBody = this.element.GetOrCreateTextBody();
            BodyProperties bodyProperties = textBody.GetFirstChild<BodyProperties>() ?? textBody.AppendChild(new BodyProperties());

            bodyProperties.LeftInset = (int)left;
            bodyProperties.TopInset = (int)top;
            bodyProperties.RightInset = (int)right;
            bodyProperties.BottomInset = (int)bottom;

            return this.result;
        }

        public OpenXmlMargin GetTextMargin()
        {
            OpenXmlElement textBody = this.element.GetTextBody();
            BodyProperties bodyProperties = textBody.GetFirstChild<BodyProperties>();

            return new OpenXmlMargin(
                bodyProperties?.LeftInset ?? (long)OpenXmlUnit.Inch(0.1),
                bodyProperties?.TopInset ?? (long)OpenXmlUnit.Inch(0.05),
                bodyProperties?.RightInset ?? (long)OpenXmlUnit.Inch(0.1),
                bodyProperties?.BottomInset ?? (long)OpenXmlUnit.Inch(0.05)
            );
        }
    }
}
