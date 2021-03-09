using System;
using DocumentFormat.OpenXml.Drawing;

namespace BinaryMesh.OpenXml.Presentations
{
    public interface ITextShape<TSelf>
    {
        TSelf SetText(string text);

        TSelf SetFont(string typeface);

        TSelf SetFontSize(int fontSize);

        TSelf SetFontColor(OpenXmlColor color);

        TSelf SetTextAnchor(TextAnchoringTypeValues anchor);

        TSelf SetTextMargin(long left, long top, long right, long bottom);

        TSelf SetIsBold(bool bold);

        TSelf SetIsItalic(bool italic);

        TSelf SetFill(OpenXmlColor color);

        TSelf SetStroke(OpenXmlColor color);
    }
}
