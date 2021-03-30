using System;
using DocumentFormat.OpenXml.Drawing;

namespace BinaryMesh.OpenXml.Presentations
{
    public interface ITextContent<out TFluent> : ITextStyle<TFluent>
    {
        TFluent SetText(string text);
    }
}
