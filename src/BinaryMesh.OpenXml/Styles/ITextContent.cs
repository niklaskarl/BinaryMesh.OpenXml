using System;
using DocumentFormat.OpenXml.Drawing;

namespace BinaryMesh.OpenXml.Styles
{
    public interface ITextContent<out TFluent> : ITextStyle<TFluent>
    {
        TFluent SetText(string text);
    }
}
