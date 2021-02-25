using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal sealed class OpenXmlShapeVisual : IShapeVisual, IOpenXmlVisual, IVisual
    {
        private readonly IOpenXmlPresentation presentation;

        private readonly Shape shape;

        public OpenXmlShapeVisual(IOpenXmlPresentation presentation, Shape shape)
        {
            this.presentation = presentation;
            this.shape = shape;
        }

        public uint Id => this.shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id;

        public string Name => this.shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name;

        public bool IsPlaceholder => this.shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape != null;

        public IShapeVisual AsShapeVisual()
        {
            return this;
        }

        public IShapeVisual SetText(string text)
        {
            if (this.shape.TextBody == null)
            {
                this.shape.TextBody = new TextBody();
            }
            else
            {
                this.shape.TextBody.RemoveAllChildren<Drawing.Paragraph>();
            }

            this.shape.TextBody.AppendChildFluent(
                new Drawing.Paragraph()
                {
    
                }
                .AppendChildFluent(
                    new Drawing.Run()
                    {
                        Text = new Drawing.Text() { Text = text }
                    }
                )
            );

            return this;
        }

        public OpenXmlElement CloneForSlide()
        {
            return this.shape.CloneNode(true);
        }
    }
}
