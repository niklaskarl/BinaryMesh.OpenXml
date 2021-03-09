using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal abstract class OpenXmlTextShapeBase<T> : ITextShape<T>
    {
        protected OpenXmlTextShapeBase()
        {
        }

        protected abstract T Self { get; }

        protected abstract OpenXmlElement GetTextBody();

        protected abstract OpenXmlElement GetOrCreateTextBody();

        protected abstract OpenXmlElement GetShapeProperties();

        protected abstract OpenXmlElement GetOrCreateShapeProperties();

        public T SetText(string text)
        {
            OpenXmlElement textBody = this.GetOrCreateTextBody();
            Paragraph paragraph = textBody.GetFirstChild<Paragraph>() ?? new Paragraph();
            Run run = paragraph.GetFirstChild<Run>() ?? new Run() { RunProperties = new RunProperties() };
            run.Text = new Text() { Text = text };

            paragraph.RemoveAllChildren<Run>();
            paragraph.AppendChild(run);

            textBody.RemoveAllChildren<Paragraph>();
            textBody.AppendChild(paragraph);

            return this.Self;
        }

        public T SetFontSize(int fontSize)
        {
            OpenXmlElement textBody = this.GetOrCreateTextBody();
            BodyProperties bodyProperties = textBody.GetFirstChild<BodyProperties>() ?? textBody.AppendChild(new BodyProperties());

            foreach (Paragraph paragraph in textBody.Elements<Paragraph>())
            {
                foreach (Run run in paragraph.Elements<Run>())
                {
                    RunProperties runProperties = run.GetFirstChild<RunProperties>() ?? run.PrependChild(new RunProperties());
                    runProperties.FontSize = fontSize * 100;
                }
            }

            return this.Self;
        }

        public T SetFont(string typeface)
        {
            OpenXmlElement textBody = this.GetOrCreateTextBody();
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

            return this.Self;
        }

        public T SetFontColor(OpenXmlColor color)
        {
            OpenXmlElement textBody = this.GetOrCreateTextBody();
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

            return this.Self;
        }

        public T SetTextAnchor(TextAnchoringTypeValues anchor)
        {
            OpenXmlElement textBody = this.GetOrCreateTextBody();
            BodyProperties bodyProperties = textBody.GetFirstChild<BodyProperties>() ?? textBody.AppendChild(new BodyProperties());
            bodyProperties.Anchor = anchor;

            return this.Self;
        }

        public T SetIsBold(bool bold)
        {
            OpenXmlElement textBody = this.GetOrCreateTextBody();
            BodyProperties bodyProperties = textBody.GetFirstChild<BodyProperties>() ?? textBody.AppendChild(new BodyProperties());

            foreach (Paragraph paragraph in textBody.Elements<Paragraph>())
            {
                foreach (Run run in paragraph.Elements<Run>())
                {
                    RunProperties runProperties = run.GetFirstChild<RunProperties>() ?? run.PrependChild(new RunProperties());
                    runProperties.Bold = bold;
                }
            }

            return this.Self;
        }

        public T SetIsItalic(bool italic)
        {
            OpenXmlElement textBody = this.GetOrCreateTextBody();
            BodyProperties bodyProperties = textBody.GetFirstChild<BodyProperties>() ?? textBody.AppendChild(new BodyProperties());

            foreach (Paragraph paragraph in textBody.Elements<Paragraph>())
            {
                foreach (Run run in paragraph.Elements<Run>())
                {
                    RunProperties runProperties = run.GetFirstChild<RunProperties>() ?? run.PrependChild(new RunProperties());
                    runProperties.Italic = italic;
                }
            }

            return this.Self;
        }

        public T SetFill(OpenXmlColor color)
        {
            OpenXmlElement shapeProperties = this.GetOrCreateShapeProperties();
            shapeProperties.RemoveAllChildren<NoFill>();
            shapeProperties.RemoveAllChildren<SolidFill>();
            shapeProperties.RemoveAllChildren<GradientFill>();
            shapeProperties.RemoveAllChildren<BlipFill>();
            shapeProperties.RemoveAllChildren<PatternFill>();
            shapeProperties.RemoveAllChildren<GroupFill>();

            shapeProperties.AppendChild(new SolidFill().AppendChildFluent(color.CreateColorElement()));

            return this.Self;
        }

        public T SetStroke(OpenXmlColor color)
        {
            OpenXmlElement shapeProperties = this.GetOrCreateShapeProperties();
            Outline outline = shapeProperties.GetFirstChild<Outline>() ?? shapeProperties.AppendChild(new Outline() { Width = 12700 });
            outline.RemoveAllChildren<NoFill>();
            outline.RemoveAllChildren<SolidFill>();
            outline.RemoveAllChildren<GradientFill>();
            outline.RemoveAllChildren<BlipFill>();
            outline.RemoveAllChildren<PatternFill>();
            outline.RemoveAllChildren<GroupFill>();

            outline.AppendChild(new SolidFill().AppendChildFluent(color.CreateColorElement()));

            return this.Self;
        }

        public T SetTextMargin(long left, long top, long right, long bottom)
        {
            OpenXmlElement textBody = this.GetOrCreateTextBody();
            BodyProperties bodyProperties = textBody.GetFirstChild<BodyProperties>() ?? textBody.AppendChild(new BodyProperties());

            bodyProperties.LeftInset = (int)left;
            bodyProperties.TopInset = (int)top;
            bodyProperties.RightInset = (int)right;
            bodyProperties.BottomInset = (int)bottom;

            return this.Self;
        }
    }
}
