using System;

using BinaryMesh.OpenXml.Styles;

namespace BinaryMesh.OpenXml.Tables
{
    public interface ITableCellStyle
    {
        IFillStyle<ITablePartStyle> Fill { get; }

        ITableCellBoderStyle Border { get; }
    }
}
