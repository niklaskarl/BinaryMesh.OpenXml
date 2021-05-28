using System;

using BinaryMesh.OpenXml.Styles;

namespace BinaryMesh.OpenXml.Tables
{
    public interface ITableCell
    {
        IVisualStyle<ITableCell> Style { get; }

        ITextContent<ITableCell> Text { get; }
    }
}
