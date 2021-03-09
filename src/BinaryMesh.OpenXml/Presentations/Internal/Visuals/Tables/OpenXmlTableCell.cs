using System;
using Drawing = DocumentFormat.OpenXml.Drawing;

using DocumentFormat.OpenXml;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal sealed class OpenXmlTableCell : OpenXmlTextShapeBase<ITableCell>, ITableCell
    {
        private readonly OpenXmlTableVisual tableVisual;

        private readonly Drawing.TableCell tableCell;

        public OpenXmlTableCell(OpenXmlTableVisual tableVisual, Drawing.TableCell tableCell)
        {
            this.tableVisual = tableVisual;
            this.tableCell = tableCell;
        }

        protected override ITableCell Self => this;

        protected override OpenXmlElement GetTextBody()
        {
            return this.tableCell.TextBody;
        }

        protected override OpenXmlElement GetOrCreateTextBody()
        {
            if (this.tableCell.TextBody == null)
            {
                this.tableCell.TextBody = new Drawing.TextBody();
            }

            return this.tableCell.TextBody;
        }

        protected override OpenXmlElement GetShapeProperties()
        {
            throw new NotSupportedException();
        }

        protected override OpenXmlElement GetOrCreateShapeProperties()
        {
            throw new NotSupportedException();
        }
    }
}
