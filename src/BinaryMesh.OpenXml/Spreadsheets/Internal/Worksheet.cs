using System;
using Packaging = DocumentFormat.OpenXml.Packaging;

namespace BinaryMesh.OpenXml.Spreadsheets.Internal
{
    internal sealed class Worksheet : IWorksheet
    {
        private readonly Packaging.WorksheetPart worksheetPart;

        public Worksheet(Packaging.WorksheetPart worksheetPart)
        {
            this.worksheetPart = worksheetPart;
        }

        public Packaging.WorksheetPart WorksheetPart => this.worksheetPart;

        public IWorksheetCells Cells => new WorksheetCells(this);

        private sealed class WorksheetCells : IWorksheetCells
        {
            private readonly Worksheet worksheet;

            public WorksheetCells(Worksheet worksheet)
            {
                this.worksheet = worksheet;
            }

            public ICell this[string reference] => Cell.TryCreateCell(this.worksheet, reference, out Cell cell) ? cell : throw new FormatException();

            public ICell this[uint column, uint row] => new Cell(this.worksheet, column, row);
        }
    }
}
