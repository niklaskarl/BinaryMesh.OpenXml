using System;
using DocumentFormat.OpenXml.Packaging;

namespace BinaryMesh.OpenXml.Spreadsheets.Internal
{
    internal sealed class OpenXmlWorksheet : IWorksheet
    {
        private readonly WorksheetPart worksheetPart;

        public OpenXmlWorksheet(WorksheetPart worksheetPart)
        {
            this.worksheetPart = worksheetPart;
        }

        public WorksheetPart WorksheetPart => this.worksheetPart;

        public IWorksheetCells Cells => new WorksheetCells(this);

        private sealed class WorksheetCells : IWorksheetCells
        {
            private readonly OpenXmlWorksheet worksheet;

            public WorksheetCells(OpenXmlWorksheet worksheet)
            {
                this.worksheet = worksheet;
            }

            public ICell this[string reference] => OpenXmlCell.TryCreateCell(this.worksheet, reference, out OpenXmlCell cell) ? cell : throw new FormatException();

            public ICell this[uint column, uint row] => new OpenXmlCell(this.worksheet, column, row);
        }
    }
}
