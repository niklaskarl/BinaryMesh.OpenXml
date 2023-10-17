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

        public IRowIterator OpenRowIterator(uint initialColumnIndex, uint initialRowIndex)
        {
            return new RowIterator(this, initialRowIndex, initialColumnIndex);
        }

        private static Row GetOrCreateInternalRow(SheetData sheetData, uint rowIndex)
        {
            Row row = sheetData.Elements<Row>()
                .SkipWhile(r => r.RowIndex - 1 < rowIndex)
                .FirstOrDefault();

            if (row == null)
            {
                row = sheetData.AppendChild(new Row() { RowIndex = rowIndex + 1 });
            }
            else if (row.RowIndex - 1 != rowIndex)
            {
                row = sheetData.InsertBefore(new Row() { RowIndex = rowIndex + 1 }, row);
            }

            return row;
        }

        private static Cell GetOrCreateInternalCell(Row row, uint columnIndex)
        {
            Cell cell = row.Elements<Cell>()
                .SkipWhile(c => !ReferenceEncoder.TryDecodeCellReference(c.CellReference, out uint cellColumnIndex, out bool isColumnFixed, out uint cellRowIndex, out bool isRowFixed) || cellColumnIndex < columnIndex)
                .FirstOrDefault();

            if (cell == null)
            {
                string reference = $"{ReferenceEncoder.EncodeColumnReference(columnIndex)}{row.RowIndex}";
                cell = row.AppendChild(new Cell() { CellReference = reference });
            }
            else if (!ReferenceEncoder.TryDecodeCellReference(cell.CellReference, out uint cellColumnIndex, out bool isColumnFixed, out uint cellRowIndex, out bool isRowFixed) || cellColumnIndex != columnIndex)
            {
                string reference = $"{ReferenceEncoder.EncodeColumnReference(columnIndex)}{row.RowIndex}";
                cell = row.InsertBefore(new Cell() { CellReference = reference }, cell);
            }

            return cell;
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

        private sealed class RowIterator : IRowIterator
        {
            private readonly OpenXmlWorksheet worksheet;

            private readonly SheetData sheetData;

            private readonly uint initialRowIndex;

            private readonly uint initialColumnIndex;

            private uint rowIndex;

            private uint columnIndex;

            private Row row;

            private Cell cell;

            public RowIterator(OpenXmlWorksheet worksheet, uint initialRowIndex, uint initialColumnIndex)
            {
                this.worksheet = worksheet;
                this.sheetData = this.worksheet.WorksheetPart.Worksheet.GetFirstChild<SheetData>();

                this.initialRowIndex = initialRowIndex;
                this.initialColumnIndex = initialColumnIndex;

                this.rowIndex = this.initialRowIndex;
                this.columnIndex = this.initialColumnIndex;

                this.row = OpenXmlWorksheet.GetOrCreateInternalRow(this.sheetData, initialRowIndex);
            }

            public uint CurrentRowIndex => this.rowIndex;

            public uint CurrentColumnIndex => this.columnIndex;

            public string CurrentReference => $"{ReferenceEncoder.EncodeColumnReference(this.columnIndex)}{this.rowIndex + 1}";

            public void NextRow()
            {
                this.rowIndex++;
                this.columnIndex = this.initialColumnIndex;
            }

            public void NextColumn()
            {
                this.columnIndex++;
            }

            public void SetValue(double value)
            {
                Cell cell = this.GetOrCreateInternalCell();
                cell.DataType = CellValues.Number;
                cell.CellValue = new CellValue(value);
            }

            public void SetValue(string value)
            {
                Cell cell = this.GetOrCreateInternalCell();
                cell.DataType = CellValues.String;
                cell.CellValue = new CellValue(value);
            }

            public void Dispose()
            {
            }

            private Cell GetOrCreateInternalCell()
            {
                if (this.row == null)
                {
                    this.row = this.sheetData.Elements<Row>().FirstOrDefault();
                    if (this.row == null)
                    {
                        this.row = this.sheetData.AppendChild(new Row() { RowIndex = this.rowIndex + 1 });
                    }
                    else if (this.row.RowIndex - 1 > this.rowIndex)
                    {
                        this.row = this.sheetData.InsertBefore(new Row() { RowIndex = this.rowIndex + 1 }, this.row);
                    }

                    this.cell = null;
                }
                else if (this.row.RowIndex - 1 < this.rowIndex)
                {
                    do
                    {
                        this.row = this.row.NextSibling<Row>();
                        this.cell = null;

                        if (this.row == null)
                        {
                            this.row = this.sheetData.AppendChild(new Row() { RowIndex = this.rowIndex + 1 });
                        }
                    } while (this.row.RowIndex - 1 < this.rowIndex);
                }

                if (this.cell == null)
                {
                    this.cell = this.row.Elements<Cell>().FirstOrDefault();
                    if (this.cell == null)
                    {
                        this.cell = this.row.AppendChild(new Cell() { CellReference = $"{ReferenceEncoder.EncodeColumnReference(this.columnIndex)}{this.rowIndex + 1}" });
                    }
                    else if (!ReferenceEncoder.TryDecodeCellReference(this.cell.CellReference, out uint oldColumnIndex, out bool _, out uint _, out bool _) || oldColumnIndex > this.columnIndex)
                    {
                        this.cell = this.row.InsertBefore(new Cell() { CellReference = $"{ReferenceEncoder.EncodeColumnReference(this.columnIndex)}{this.rowIndex + 1}" }, this.cell);
                    }
                }
                else if (!ReferenceEncoder.TryDecodeCellReference(this.cell.CellReference, out uint oldColumnIndex, out bool _, out uint _, out bool _) || oldColumnIndex < this.columnIndex)
                {
                    do
                    {
                        this.cell = this.cell.NextSibling<Cell>();

                        if (this.cell == null)
                        {
                            this.cell = this.row.AppendChild(new Cell() { CellReference = $"{ReferenceEncoder.EncodeColumnReference(this.columnIndex)}{this.rowIndex + 1}" });
                        }
                    } while (!ReferenceEncoder.TryDecodeCellReference(this.cell.CellReference, out oldColumnIndex, out bool _, out uint _, out bool _) || oldColumnIndex < this.columnIndex);
                }

                return this.cell;
            }
        }
    }
}
