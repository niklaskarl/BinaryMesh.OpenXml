using System;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;

namespace BinaryMesh.OpenXml.Spreadsheets.Internal
{
    internal sealed class OpenXmlCell : ICell
    {
        private readonly OpenXmlWorksheet worksheet;

        private readonly uint column;

        private readonly bool isColumnFixed;

        private readonly uint row;

        private readonly bool isRowFixed;

        public OpenXmlCell(OpenXmlWorksheet worksheet, uint column, uint row)
        {
            this.worksheet = worksheet;
            this.column = column;
            this.isColumnFixed = false;
            this.row = row;
            this.isRowFixed = false;
        }

        public OpenXmlCell(OpenXmlWorksheet worksheet, uint column, bool isColumnFixed, uint row, bool isRowFixed)
        {
            this.worksheet = worksheet;
            this.column = column;
            this.isColumnFixed = isColumnFixed;
            this.row = row;
            this.isRowFixed = isRowFixed;
        }

        public uint Column => this.column;

        public bool IsColumnFixed => this.isColumnFixed;

        public uint Row => this.row;

        public bool IsRowFixed => this.isRowFixed;

        public string Reference => $"{(this.isColumnFixed ? "$" : "")}{GetColumnCharsFromIndex(this.column)}{(this.isRowFixed ? "$" : "")}{this.row + 1}";

        public static bool TryCreateCell(OpenXmlWorksheet worksheet, string reference, out OpenXmlCell cell)
        {
            if (TryDecodeCellReference(reference, out uint column, out bool isColumnFixed, out uint row, out bool isRowFixed))
            {
                cell = new OpenXmlCell(worksheet, column, isColumnFixed, row, isRowFixed);
                return true;
            }
            else
            {
                cell = null;
                return false;
            }
        }

        public ICell SetValue(double value)
        {
            Cell cell = this.GetOrCreateInternalCell();
            cell.DataType = CellValues.Number;
            cell.CellValue = new CellValue(value);

            return this;
        }

        public ICell SetValue(string value)
        {
            Cell cell = this.GetOrCreateInternalCell();
            cell.DataType = CellValues.String;
            cell.CellValue = new CellValue(value);

            return this;
        }

        private Cell GetInternalCell()
        {
            SheetData sheetData = this.worksheet.WorksheetPart.Worksheet.GetFirstChild<SheetData>();
            Row row = sheetData.Elements<Row>()
                .SkipWhile(r => r.RowIndex - 1 < this.row)
                .FirstOrDefault();

            if (row != null && row.RowIndex - 1 == this.row)
            {
                Cell cell = row.Elements<Cell>()
                    .SkipWhile(c => !TryDecodeCellReference(c.CellReference, out uint columnIndex, out bool isColumnFixed, out uint rowIndex, out bool isRowFixed) || columnIndex < this.column)
                    .FirstOrDefault();

                {
                    if (cell != null && TryDecodeCellReference(cell.CellReference, out uint columnIndex, out bool isColumnFixed, out uint rowIndex, out bool isRowFixed) && columnIndex == this.column)
                    {
                        return cell;
                    }
                }
            }

            return null;
        }

        private Cell GetOrCreateInternalCell()
        {
            SheetData sheetData = this.worksheet.WorksheetPart.Worksheet.GetFirstChild<SheetData>();
            Row row = sheetData.Elements<Row>()
                .SkipWhile(r => r.RowIndex - 1 < this.row)
                .FirstOrDefault();

            Cell cell;
            if (row != null && row.RowIndex - 1 == this.row)
            {
                cell = row.Elements<Cell>()
                    .SkipWhile(c => !TryDecodeCellReference(c.CellReference, out uint columnIndex, out bool isColumnFixed, out uint rowIndex, out bool isRowFixed) || columnIndex < this.column)
                    .FirstOrDefault();

                {
                    if (cell == null)
                    {
                        cell = row.AppendChild(new Cell() { CellReference = $"{GetColumnCharsFromIndex(this.column)}{this.row + 1}" });
                    }
                    else if (!TryDecodeCellReference(cell.CellReference, out uint columnIndex, out bool isColumnFixed, out uint rowIndex, out bool isRowFixed) || columnIndex != this.column)
                    {
                        cell = row.InsertBefore(new Cell() { CellReference = $"{GetColumnCharsFromIndex(this.column)}{this.row + 1}" }, cell);
                    }
                }
            }
            else
            {
                if (row != null)
                {
                    row = sheetData.InsertBefore(new Row() { RowIndex = this.row + 1 }, row);
                }
                else
                {
                    row = sheetData.AppendChild(new Row() { RowIndex = this.row + 1 });
                }

                cell = row.AppendChild(new Cell() { CellReference = $"{GetColumnCharsFromIndex(this.column)}{this.row + 1}" });
            }

            return cell;
        }

        private static string GetColumnCharsFromIndex(uint column)
        {
            string result = ((char)('A' + (column) % ('Z' - 'A' + 1))).ToString();
            while((column = (column) / ('Z' - 'A' + 1)) > 0)
            {
                result = ((char)('A' + (column) % ('Z' - 'A' + 1))) + result;
            }

            return result;
        }

        private static bool TryDecodeCellReference(string reference, out uint column, out bool isColumnFixed, out uint row, out bool isRowFixed)
        {
            column = 0;
            isColumnFixed = false;
            row = 0;
            isRowFixed = false;

            int i = 0;
            if (reference[i] == '$')
            {
                isColumnFixed = true;
                ++i;
            }

            if (!(i < reference.Length && ((reference[i] >= 'A' && reference[i] <= 'Z') || (reference[i] >= 'a' && reference[i] <= 'z'))))
            {
                return false;
            }

            while (i < reference.Length && ((reference[i] >= 'A' && reference[i] <= 'Z') || (reference[i] >= 'a' && reference[i] <= 'z')))
            {
                if (reference[i] < 'a')
                {
                    column *= 'Z' - 'A' + 1;
                    column += ((uint)(reference[i] - 'A') + 1u);
                }
                else
                {
                    column *= 'z' - 'a' + 1;
                    column += ((uint)(reference[i] - 'a') + 1u);
                }

                ++i;
            }

            if (i == reference.Length)
            {
                return false;
            }

            if (reference[i] == '$')
            {
                isRowFixed = true;
                ++i;
            }

            if (i == reference.Length)
            {
                return false;
            }

            if (!uint.TryParse(reference.Substring(i), out row))
            {
                return false;
            }

            column -= 1;
            row -= 1;

            return true;
        }
    }
}
