using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

using BinaryMesh.OpenXml.Helpers;
using BinaryMesh.OpenXml.Spreadsheets.Helpers;

namespace BinaryMesh.OpenXml.Spreadsheets.Internal
{
    internal class OpenXmlWorkbook : IWorkbook
    {
        private readonly WorkbookPart workbookPart;

        public OpenXmlWorkbook(WorkbookPart workbookPart)
        {
            this.workbookPart = workbookPart;
            if (this.workbookPart.Workbook == null)
            {
                this.workbookPart.Workbook = new Workbook();
            }

            if (this.workbookPart.Workbook.Sheets == null)
            {
                this.workbookPart.Workbook.Sheets = new Sheets();
            }
        }

        public WorkbookPart WorkbookPart => this.workbookPart;

        public OpenXmlWorksheet GetWorksheet(string name)
        {
            Sheet sheet = this.workbookPart.Workbook.Sheets.Elements<Sheet>().FirstOrDefault(s => s.Name == name);
            if (sheet != null)
            {
                return new OpenXmlWorksheet(this, this.workbookPart.GetPartById(sheet.Id.Value) as WorksheetPart);
            }

            return null;
        }

        public KeyedReadOnlyList<string, IWorksheet> Worksheets => new EnumerableKeyedList<Sheet, string, IWorksheet>(
            this.workbookPart.Workbook.Sheets.Elements<Sheet>(),
            sheet => sheet.Name,
            sheet => new OpenXmlWorksheet(this, this.workbookPart.GetPartById(sheet.Id.Value) as WorksheetPart)
        );

        public IWorksheet AppendWorksheet(string name)
        {
            Sheet refSheet = this.workbookPart.Workbook.Sheets.Elements<Sheet>()
                .LastOrDefault();

            return this.InsertWorksheetAfter(name, refSheet);
        }

        public IWorksheet InsertWorksheet(string name, int index)
        {
            Sheet refSheet = this.workbookPart.Workbook.Sheets.Elements<Sheet>()
                .Skip(index).LastOrDefault();

            return this.InsertWorksheetAfter(name, refSheet);
        }

        public IWorksheet PrependWorksheet(string name)
        {
            return this.InsertWorksheetAfter(name, null);
        }

        private IWorksheet InsertWorksheetAfter(string name, Sheet refSheet)
        {
            WorksheetPart worksheetPart = this.workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            uint sheetId = (refSheet?.SheetId ?? 0u) + 1u;
            Sheet sheet = this.workbookPart.Workbook.Sheets.InsertAfter(
                new Sheet() { Id = this.workbookPart.GetIdOfPart(worksheetPart), SheetId = sheetId, Name = name },
                refSheet
            );

            sheet = sheet.NextSibling<Sheet>();
            ++sheetId;
            while (sheet != null)
            {
                sheet.SheetId = sheetId;

                sheet = sheet.NextSibling<Sheet>();
                ++sheetId;
            }

            return new OpenXmlWorksheet(this, worksheetPart);
        }

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

            if (worksheetName == null)
            {
                throw new ArgumentException("No worksheet specified");
            }

            if (!this.Worksheets.TryGetValue(worksheetName, out IWorksheet worksheet))
            {
                throw new ArgumentException("Invalid worksheet name.");
            }

            return new OpenXmlRange(
                worksheet as OpenXmlWorksheet,
                startColumn, isStartColumnFixed,
                startRow, isStartRowFixed,
                endColumn, isEndColumnFixed,
                endRow, isEndRowFixed
            );
        }
    }
}
