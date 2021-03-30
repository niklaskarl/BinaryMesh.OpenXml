using System;
using DocumentFormat.OpenXml.Drawing;

namespace BinaryMesh.OpenXml.Presentations
{
    public interface ITextStyle<out TFluent>
    {
        TFluent SetFont(string typeface);

        TFluent SetFontSize(double fontSize);

        TFluent SetFontColor(OpenXmlColor color);

        TFluent SetTextAlign(TextAlignmentTypeValues align);

        TFluent SetTextAnchor(TextAnchoringTypeValues anchor);

        TFluent SetTextMargin(long left, long top, long right, long bottom);

        TFluent SetIsBold(bool bold);

        TFluent SetIsItalic(bool italic);
    }
}
