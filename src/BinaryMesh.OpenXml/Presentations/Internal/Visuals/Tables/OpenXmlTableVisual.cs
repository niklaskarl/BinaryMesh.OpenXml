using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;

using BinaryMesh.OpenXml.Helpers;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal sealed class OpenXmlTableVisual : OpenXmlGraphicFrameVisual, ITableVisual
    {
        private readonly Drawing.Table table;

        public OpenXmlTableVisual(IOpenXmlVisualContainer container, GraphicFrame graphicFrame) :
            base(container, graphicFrame)
        {
            this.table = this.graphicFrame
                .GetFirstChild<Drawing.Graphic>()
                .GetFirstChild<Drawing.GraphicData>()
                .GetFirstChild<Drawing.Table>();
        }

        public ITableCellCollection Cells => new TableCellCollection(this);

        public IReadOnlyList<ITableColumn> Columns => new EnumerableList<Drawing.GridColumn, ITableColumn>(this.table.GetFirstChild<Drawing.TableGrid>().Elements<Drawing.GridColumn>(), column => new TableColumn(this, column));

        public IReadOnlyList<ITableRow> Rows => new EnumerableList<Drawing.TableRow, ITableRow>(this.table.Elements<Drawing.TableRow>(), row => new TableRow(this, row));

        public ITableColumn AppendColumn(long width)
        {
            Drawing.GridColumn gridColumn = this.table.GetFirstChild<Drawing.TableGrid>().AppendChild(
                new Drawing.GridColumn()
                {
                    Width = width
                }
            );

            foreach (Drawing.TableRow tableRow in this.table.Elements<Drawing.TableRow>())
            {
                tableRow.AppendChild(
                    new Drawing.TableCell()
                    {
                        TableCellProperties = new Drawing.TableCellProperties(),
                        TextBody = new Drawing.TextBody()
                        {
                            BodyProperties = new Drawing.BodyProperties(),
                            ListStyle = new Drawing.ListStyle()
                        }
                            .AppendChildFluent(new Drawing.Paragraph().AppendChildFluent(new Drawing.Run().AppendChildFluent(new Drawing.Text())))
                    }
                );
            }

            return new TableColumn(this, gridColumn);
        }

        public ITableRow AppendRow(long height)
        {
            Drawing.TableRow tableRow = this.table.AppendChild(
                new Drawing.TableRow()
                {
                    Height = height
                }
            );

            foreach (Drawing.GridColumn gridColumn in this.table.GetFirstChild<Drawing.TableGrid>().Elements<Drawing.GridColumn>())
            {
                tableRow.AppendChild(
                    new Drawing.TableCell()
                    {
                        TableCellProperties = new Drawing.TableCellProperties(),
                        TextBody = new Drawing.TextBody()
                        {
                            BodyProperties = new Drawing.BodyProperties(),
                            ListStyle = new Drawing.ListStyle()
                        }
                            .AppendChildFluent(new Drawing.Paragraph().AppendChildFluent(new Drawing.Run().AppendChildFluent(new Drawing.Text())))
                    }
                );
            }

            return new TableRow(this, tableRow);
        }

        ITableVisual ITableVisual.SetOffset(long x, long y)
        {
            this.SetOffset(x, y);
            return this;
        }

        ITableVisual ITableVisual.SetExtents(long width, long height)
        {
            this.SetExtents(width, height);
            return this;
        }

        private sealed class TableColumn : ITableColumn
        {
            private readonly OpenXmlTableVisual tableVisual;

            private readonly Drawing.GridColumn gridColumn;

            public TableColumn(OpenXmlTableVisual tableVisual, Drawing.GridColumn gridColumn)
            {
                this.tableVisual = tableVisual;
                this.gridColumn = gridColumn;
            }
        }

        private sealed class TableRow : ITableRow
        {
            private readonly OpenXmlTableVisual tableVisual;

            private readonly Drawing.TableRow tableRow;

            public TableRow(OpenXmlTableVisual tableVisual, Drawing.TableRow tableRow)
            {
                this.tableVisual = tableVisual;
                this.tableRow = tableRow;
            }
        }

        private sealed class TableCellCollection : ITableCellCollection
        {
            private readonly OpenXmlTableVisual tableVisual;

            public TableCellCollection(OpenXmlTableVisual tableVisual)
            {
                this.tableVisual = tableVisual;
            }

            public ITableCell this[int column, int row] =>
                new OpenXmlTableCell(
                    this.tableVisual,
                    this.tableVisual.table.Elements<Drawing.TableRow>().ElementAt(row).Elements<Drawing.TableCell>().ElementAt(column)
                );
        }
    }
}
