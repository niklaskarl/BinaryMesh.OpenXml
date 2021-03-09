using System;
using System.Linq;
using BinaryMesh.OpenXml.Spreadsheets.Helpers;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace BinaryMesh.OpenXml.Spreadsheets.Internal
{
    internal sealed class OpenXmlWorksheet : IWorksheet
    {
        private readonly OpenXmlWorkbook workbook;

        private readonly WorksheetPart worksheetPart;

        public OpenXmlWorksheet(OpenXmlWorkbook workbook, WorksheetPart worksheetPart)
        {
            this.workbook = workbook;
            this.worksheetPart = worksheetPart;
        }

        public WorksheetPart WorksheetPart => this.worksheetPart;

        public string Name
        {
            get
            {
                WorkbookPart workbookPart = this.workbook.WorkbookPart;
                string id = workbookPart.GetIdOfPart(this.worksheetPart);
                Sheet sheet = workbookPart.Workbook.Sheets.Elements<Sheet>().FirstOrDefault(s => s.Id == id);
                return sheet?.Name;
            }
        }

        public IWorksheetCells Cells => new WorksheetCells(this);

        public IRange GetRange(string formula)
        {
            bool result = ReferenceEncoder.TryDecodeRangeReference(
                formula,
                out string worksheetName,
                out uint? startColumn, out bool isStartColumnFixed,
                out uint? startRow, out bool isStartRowFixed,
                out uint? endColumn, out bool isEndColumnFixed,
                out uint? endRow, out bool isEndRowFixed
            );

            if (!result)
            {
                throw new FormatException();
            }

            if (worksheetName != null && worksheetName != this.Name)
            {
                throw new ArgumentException("Referenced a different worksheet");
            }

            return new OpenXmlRange(
                this,
                startColumn, isStartColumnFixed,
                startRow, isStartRowFixed,
                endColumn, isEndColumnFixed,
                endRow, isEndRowFixed
            );
        }

        public IRange GetRange(uint startColumn, uint startRow, uint endColumn, uint endRow)
        {
            return new OpenXmlRange(
                this,
                startColumn, false,
                startRow, false,
                endColumn, false,
                endRow, false
            );
        }

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
