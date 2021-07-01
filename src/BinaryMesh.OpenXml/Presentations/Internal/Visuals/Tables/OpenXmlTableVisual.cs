using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;

using BinaryMesh.OpenXml.Helpers;
using BinaryMesh.OpenXml.Styles;
using BinaryMesh.OpenXml.Styles.Internal;
using BinaryMesh.OpenXml.Styles.Internal.Mixins;
using BinaryMesh.OpenXml.Tables;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal sealed class OpenXmlTableVisual : OpenXmlGraphicFrameVisual<ITableVisual>, IVisualTransform<ITableVisual>, ITableVisual
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

        public IVisualTransform<ITableVisual> Transform => this;

        protected override ITableVisual Self => this;

        public ITableVisual SetStyle(ITableStyle style)
        {
            string id = style.Id;
            this.GetTableStyleId(true).Text = id;

            return this;
        }

        public ITableVisual SetHasFirstRow(bool value)
        {
            this.GetTableProperties(true).FirstRow = value;

            return this;
        }

        public ITableVisual SetHasFirstColumn(bool value)
        {
            this.GetTableProperties(true).FirstColumn = value;

            return this;
        }

        public ITableVisual SetHasLastRow(bool value)
        {
            this.GetTableProperties(true).LastRow = value;

            return this;
        }

        public ITableVisual SetHasLastColumn(bool value)
        {
            this.GetTableProperties(true).LastColumn = value;

            return this;
        }

        public ITableVisual SetHasBandRow(bool value)
        {
            this.GetTableProperties(true).BandRow = value;

            return this;
        }

        public ITableVisual SetHasBandColumn(bool value)
        {
            this.GetTableProperties(true).BandColumn = value;

            return this;
        }

        public ITableVisual SetIsRightToLeft(bool value)
        {
            this.GetTableProperties(true).RightToLeft = value;

            return this;
        }

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

        private Drawing.TableProperties GetTableProperties(bool create)
        {
            Drawing.TableProperties result = this.table.TableProperties;
            if (result == null && create)
            {
                result = new Drawing.TableProperties();
                this.table.TableProperties = result;
            }

            return result;
        }

        private Drawing.TableStyleId GetTableStyleId(bool create)
        {
            Drawing.TableProperties tableProperties = GetTableProperties(create);
            Drawing.TableStyleId result = tableProperties.GetFirstChild<Drawing.TableStyleId>();
            if (result == null && create)
            {
                result = tableProperties.AppendChild(new Drawing.TableStyleId());
            }

            return result;
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
