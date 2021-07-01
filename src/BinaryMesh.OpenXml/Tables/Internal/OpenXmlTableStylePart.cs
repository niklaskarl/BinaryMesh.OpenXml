using System;
using DocumentFormat.OpenXml.Drawing;

using BinaryMesh.OpenXml.Shared;
using BinaryMesh.OpenXml.Styles;

namespace BinaryMesh.OpenXml.Tables.Internal
{
    internal class OpenXmlTablePartStyle : ITablePartStyle
    {
        private readonly ElementGenerator<TablePartStyleType> tablePartStyleTypeGenerator;

        public OpenXmlTablePartStyle(ElementGenerator<TablePartStyleType> tablePartStyleTypeGenerator)
        {
            this.tablePartStyleTypeGenerator = tablePartStyleTypeGenerator;
        }

        public ITableCellStyle Style => new OpenXmlTableCellStyle(this, this.GetTableCellStyle);

        public ITableCellTextStyle Text => new OpenXmlTableCellTextStyle(this, this.GetTableCellTextStyle);

        private TableCellStyle GetTableCellStyle(bool create)
        {
            TablePartStyleType tablePartStyleType = this.tablePartStyleTypeGenerator(true);
            TableCellStyle result = tablePartStyleType.TableCellStyle;
            if (result == null && create)
            {
                result = new TableCellStyle();
                tablePartStyleType.TableCellStyle = result;
            }

            return result;
        }

        private TableCellTextStyle GetTableCellTextStyle(bool create)
        {
            TablePartStyleType tablePartStyleType = this.tablePartStyleTypeGenerator(true);
            TableCellTextStyle result = tablePartStyleType.TableCellTextStyle;
            if (result == null && create)
            {
                result = new TableCellTextStyle();
                tablePartStyleType.TableCellTextStyle = result;
            }

            return result;
        }
    }
}
