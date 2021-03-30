using System;
using Drawing = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml;

using BinaryMesh.OpenXml.Presentations.Internal.Mixins;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal sealed class OpenXmlTableCell : IOpenXmlTextElement, IOpenXmlShapeElement, ITableCell
    {
        private readonly OpenXmlTableVisual tableVisual;

        private readonly Drawing.TableCell tableCell;

        public OpenXmlTableCell(OpenXmlTableVisual tableVisual, Drawing.TableCell tableCell)
        {
            this.tableVisual = tableVisual;
            this.tableCell = tableCell;
        }

        public IVisualStyle<ITableCell> Style => new OpenXmlVisualStyle<OpenXmlTableCell, ITableCell>(this);

        public ITextContent<ITableCell> Text => new OpenXmlTextContent<OpenXmlTableCell, ITableCell>(this);

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
    }
}
