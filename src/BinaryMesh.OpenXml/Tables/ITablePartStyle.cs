using System;

using BinaryMesh.OpenXml.Styles;

namespace BinaryMesh.OpenXml.Tables
{
    public interface ITablePartStyle
    {
        ITableCellStyle Style { get; }

        ITableCellTextStyle Text { get; }
    }
}
