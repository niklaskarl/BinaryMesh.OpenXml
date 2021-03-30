using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;

namespace BinaryMesh.OpenXml.Presentations.Internal.Mixins
{
    internal class OpenXmlTextStyle<TElement, TFluent> : ITextStyle<TFluent>
        where TElement : IOpenXmlTextElement, TFluent
    {
        protected readonly TElement element;

        public OpenXmlTextStyle(TElement element)
        {
            this.element = element;
        }

        public TFluent SetFontSize(int fontSize)
        {
            OpenXmlElement textBody = this.element.GetOrCreateTextBody();
            BodyProperties bodyProperties = textBody.GetFirstChild<BodyProperties>() ?? textBody.AppendChild(new BodyProperties());

            foreach (Paragraph paragraph in textBody.Elements<Paragraph>())
            {
                foreach (Run run in paragraph.Elements<Run>())
                {
                    RunProperties runProperties = run.GetFirstChild<RunProperties>() ?? run.PrependChild(new RunProperties());
                    runProperties.FontSize = fontSize * 100;
                }
            }

            return this.element;
        }

        public TFluent SetFont(string typeface)
        {
            OpenXmlElement textBody = this.element.GetOrCreateTextBody();
            BodyProperties bodyProperties = textBody.GetFirstChild<BodyProperties>() ?? textBody.AppendChild(new BodyProperties());

            foreach (Paragraph paragraph in textBody.Elements<Paragraph>())
            {
                foreach (Run run in paragraph.Elements<Run>())
                {
                    RunProperties runProperties = run.GetFirstChild<RunProperties>() ?? run.PrependChild(new RunProperties());
                    runProperties.RemoveAllChildren<LatinFont>();
                    runProperties.RemoveAllChildren<ComplexScriptFont>();

                    runProperties.AppendChild(new LatinFont() { Typeface = typeface });
                    runProperties.AppendChild(new ComplexScriptFont() { Typeface = typeface });
                }
            }

            return this.element;
        }

        public TFluent SetFontColor(OpenXmlColor color)
        {
            OpenXmlElement textBody = this.element.GetOrCreateTextBody();
            BodyProperties bodyProperties = textBody.GetFirstChild<BodyProperties>() ?? textBody.AppendChild(new BodyProperties());

            foreach (Paragraph paragraph in textBody.Elements<Paragraph>())
            {
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

            return this.element;
        }

        public TFluent SetTextAnchor(TextAnchoringTypeValues anchor)
        {
            OpenXmlElement textBody = this.element.GetOrCreateTextBody();
            BodyProperties bodyProperties = textBody.GetFirstChild<BodyProperties>() ?? textBody.AppendChild(new BodyProperties());
            bodyProperties.Anchor = anchor;

            return this.element;
        }

        public TFluent SetIsBold(bool bold)
        {
            OpenXmlElement textBody = this.element.GetOrCreateTextBody();
            BodyProperties bodyProperties = textBody.GetFirstChild<BodyProperties>() ?? textBody.AppendChild(new BodyProperties());

            foreach (Paragraph paragraph in textBody.Elements<Paragraph>())
            {
                foreach (Run run in paragraph.Elements<Run>())
                {
                    RunProperties runProperties = run.GetFirstChild<RunProperties>() ?? run.PrependChild(new RunProperties());
                    runProperties.Bold = bold;
                }
            }

            return this.element;
        }

        public TFluent SetIsItalic(bool italic)
        {
            OpenXmlElement textBody = this.element.GetOrCreateTextBody();
            BodyProperties bodyProperties = textBody.GetFirstChild<BodyProperties>() ?? textBody.AppendChild(new BodyProperties());

            foreach (Paragraph paragraph in textBody.Elements<Paragraph>())
            {
                foreach (Run run in paragraph.Elements<Run>())
                {
                    RunProperties runProperties = run.GetFirstChild<RunProperties>() ?? run.PrependChild(new RunProperties());
                    runProperties.Italic = italic;
                }
            }

            return this.element;
        }

        public TFluent SetTextMargin(long left, long top, long right, long bottom)
        {
            OpenXmlElement textBody = this.element.GetOrCreateTextBody();
            BodyProperties bodyProperties = textBody.GetFirstChild<BodyProperties>() ?? textBody.AppendChild(new BodyProperties());

            bodyProperties.LeftInset = (int)left;
            bodyProperties.TopInset = (int)top;
            bodyProperties.RightInset = (int)right;
            bodyProperties.BottomInset = (int)bottom;

            return this.element;
        }
    }
}
