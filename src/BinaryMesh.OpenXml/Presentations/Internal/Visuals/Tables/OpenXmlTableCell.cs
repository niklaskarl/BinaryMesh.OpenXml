using System;
using DocumentFormat.OpenXml;
using Drawing = DocumentFormat.OpenXml.Drawing;
using SixLabors.Fonts;

using BinaryMesh.OpenXml.Internal;
using BinaryMesh.OpenXml.Styles;
using BinaryMesh.OpenXml.Styles.Internal;
using BinaryMesh.OpenXml.Styles.Internal.Mixins;
using BinaryMesh.OpenXml.Tables;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal sealed class OpenXmlTableCell : IOpenXmlTextElement, IOpenXmlShapeElement, ITableCell
    {
        private readonly OpenXmlTableVisual tableVisual;

        private readonly Drawing.TableRow tableRow;

        private readonly Drawing.TableCell tableCell;

        public OpenXmlTableCell(OpenXmlTableVisual tableVisual, Drawing.TableRow tableRow, Drawing.TableCell tableCell)
        {
            this.tableVisual = tableVisual;
            this.tableRow = tableRow;
            this.tableCell = tableCell;
        }

        public IVisualStyle<ITableCell> Style => this.InternalStyle;

        public ITextContent<ITableCell> Text => this.InternalText;

        public OpenXmlVisualStyle<OpenXmlTableCell, ITableCell> InternalStyle => new OpenXmlVisualStyle<OpenXmlTableCell, ITableCell>(this, this);

        internal OpenXmlTextContent<OpenXmlTableCell, ITableCell> InternalText => new OpenXmlTextContent<OpenXmlTableCell, ITableCell>(this, this);

        public OpenXmlElement GetTextBody()
        {
            return this.tableCell.TextBody;
        }

        public OpenXmlElement GetOrCreateTextBody()
        {
            if (this.tableCell.TextBody == null)
            {
                this.tableCell.TextBody = new Drawing.TextBody();
            }

            return this.tableCell.TextBody;
        }

        public OpenXmlElement GetShapeProperties()
        {
            return this.tableCell.TableCellProperties;
        }

        public OpenXmlElement GetOrCreateShapeProperties()
        {
            if (this.tableCell.TableCellProperties == null)
            {
                this.tableCell.TableCellProperties = new Drawing.TableCellProperties();
            }

            return this.tableCell.TableCellProperties;
        }

        public OpenXmlSize Measure()
        {
            IOpenXmlTheme theme = this.tableVisual.Container.Document.Theme;
            IOpenXmlTextStyle defaultTextStyle = this.tableVisual.Container.Document.DefaultTextStyle;

            int column = this.GetColumnIndex();
            OpenXmlUnit width = (this.tableVisual.Columns[column] as OpenXmlTableVisual.TableColumn).MeasureWidth();
            return this.InternalText.MeasureText(defaultTextStyle, theme, width);
        }

        private int GetColumnIndex()
        {
            int i = 0;
            foreach (Drawing.TableCell cell in this.tableRow.Elements<Drawing.TableCell>())
            {
                if (this.tableCell == cell)
                {
                    return i;
                }

                ++i;
            }

            return -1;
        }
    }
}
