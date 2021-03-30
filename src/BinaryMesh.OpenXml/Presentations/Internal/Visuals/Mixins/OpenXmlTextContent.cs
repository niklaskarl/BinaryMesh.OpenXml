using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;

namespace BinaryMesh.OpenXml.Presentations.Internal.Mixins
{
    internal class OpenXmlTextContent<TElement, TFluent> : OpenXmlTextStyle<TElement, TFluent>, ITextContent<TFluent>, ITextStyle<TFluent>
        where TElement : IOpenXmlTextElement, TFluent
    {
        public OpenXmlTextContent(TElement element)
            : base(element)
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

            return this.element;
        }
    }
}
