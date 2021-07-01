using System;
using System.Collections.Generic;
using BinaryMesh.OpenXml.Helpers;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;

namespace BinaryMesh.OpenXml.Tables.Internal
{
    internal class OpenXmlTableStyle : ITableStyle
    {
        private readonly TableStyleEntry tableStyleEntry;

        public OpenXmlTableStyle(TableStyleEntry tableStyleEntry)
        {
            this.tableStyleEntry = tableStyleEntry;
        }

        public string Id => this.tableStyleEntry.StyleId;

        public ITablePartStyle WholeTablePart => new OpenXmlTablePartStyle(craeate => this.tableStyleEntry.WholeTable ?? (craeate ? (this.tableStyleEntry.WholeTable = new WholeTable()) : null));

        public ITablePartStyle Row => new OpenXmlTablePartStyle(create => this.tableStyleEntry.Band1Horizontal ?? (create ? (this.tableStyleEntry.Band1Horizontal = new Band1Horizontal()) : null));

        public ITablePartStyle AlternatingRow => new OpenXmlTablePartStyle(create => this.tableStyleEntry.Band2Horizontal ?? (create ? (this.tableStyleEntry.Band2Horizontal = new Band2Horizontal()) : null));

        public ITablePartStyle Column => new OpenXmlTablePartStyle(create => this.tableStyleEntry.Band1Vertical ?? (create ? (this.tableStyleEntry.Band1Vertical = new Band1Vertical()) : null));

        public ITablePartStyle AlternatingColumn => new OpenXmlTablePartStyle(create => this.tableStyleEntry.Band2Vertical ?? (create ? (this.tableStyleEntry.Band2Vertical = new Band2Vertical()) : null));

        public ITablePartStyle LastColumn => new OpenXmlTablePartStyle(create => this.tableStyleEntry.LastColumn ?? (create ? (this.tableStyleEntry.LastColumn = new LastColumn()) : null));

        public ITablePartStyle FirstColumn => new OpenXmlTablePartStyle(create => this.tableStyleEntry.FirstColumn ?? (create ? (this.tableStyleEntry.FirstColumn = new FirstColumn()) : null));

        public ITablePartStyle LastRow => new OpenXmlTablePartStyle(create => this.tableStyleEntry.LastRow ?? (create ? (this.tableStyleEntry.LastRow = new LastRow()) : null));

        public ITablePartStyle FirstRow => new OpenXmlTablePartStyle(create => this.tableStyleEntry.FirstRow ?? (create ? (this.tableStyleEntry.FirstRow = new FirstRow()) : null));

        public static TableStyleEntry InitializeTableStyle()
        {
            return new TableStyleEntry()
            {
                WholeTable = new WholeTable() { TableCellStyle = new TableCellStyle() { TableCellBorders = new TableCellBorders() } },
                Band1Horizontal = new Band1Horizontal() { TableCellStyle = new TableCellStyle() { TableCellBorders = new TableCellBorders() } },
                Band2Horizontal = new Band2Horizontal() { TableCellStyle = new TableCellStyle() { TableCellBorders = new TableCellBorders() } },
                Band1Vertical = new Band1Vertical() { TableCellStyle = new TableCellStyle() { TableCellBorders = new TableCellBorders() } },
                Band2Vertical = new Band2Vertical() { TableCellStyle = new TableCellStyle() { TableCellBorders = new TableCellBorders() } },
                LastColumn = new LastColumn() { TableCellStyle = new TableCellStyle() { TableCellBorders = new TableCellBorders() } },
                FirstColumn = new FirstColumn() { TableCellStyle = new TableCellStyle() { TableCellBorders = new TableCellBorders() } },
                LastRow = new LastRow() { TableCellStyle = new TableCellStyle() { TableCellBorders = new TableCellBorders() } },
                FirstRow = new FirstRow() { TableCellStyle = new TableCellStyle() { TableCellBorders = new TableCellBorders() } }
            };
        }
    }
}