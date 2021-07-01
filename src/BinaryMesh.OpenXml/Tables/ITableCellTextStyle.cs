using System;

using BinaryMesh.OpenXml.Styles;

namespace BinaryMesh.OpenXml.Tables
{
    public interface ITableCellTextStyle
    {
        ITablePartStyle SetFont(string typeface);

        ITablePartStyle SetFont(OpenXmlFontRef fontRef);

        ITablePartStyle SetFontColor(OpenXmlColor color);

        ITablePartStyle SetIsBold(bool bold);

        ITablePartStyle SetIsItalic(bool italic);
    }
}
