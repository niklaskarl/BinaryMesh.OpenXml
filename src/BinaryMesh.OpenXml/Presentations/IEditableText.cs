using System;
using DocumentFormat.OpenXml.Drawing;

namespace BinaryMesh.OpenXml.Presentations
{
    public interface IEditableText<TSelf> : IStyleableText<TSelf>
    {
        TSelf SetText(string text);
    }
}
