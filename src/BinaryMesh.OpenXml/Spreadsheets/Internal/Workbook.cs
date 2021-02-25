using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using Packaging = DocumentFormat.OpenXml.Packaging;
using Spreadsheet = DocumentFormat.OpenXml.Spreadsheet;

using BinaryMesh.OpenXml.Helpers;

namespace BinaryMesh.OpenXml.Spreadsheets.Internal
{
    internal class Workbook : IWorkbook
    {
        private readonly Packaging.WorkbookPart workbookPart;

        public Workbook(Packaging.WorkbookPart workbookPart)
        {
            this.workbookPart = workbookPart;
            if (this.workbookPart.Workbook == null)
            {
                this.workbookPart.Workbook = new Spreadsheet.Workbook();
            }

            if (this.workbookPart.Workbook.Sheets == null)
            {
                this.workbookPart.Workbook.Sheets = new Spreadsheet.Sheets();
            }
        }

        public KeyedReadOnlyList<string, IWorksheet> Worksheets => new EnumerableKeyedList<Spreadsheet.Sheet, string, IWorksheet>(
            this.workbookPart.Workbook.Sheets.Elements<Spreadsheet.Sheet>(),
            sheet => sheet.Name,
            sheet => new Worksheet(this.workbookPart.GetPartById(sheet.Id.Value) as Packaging.WorksheetPart)
        );

        public IWorksheet AppendWorksheet(string name)
        {
            Spreadsheet.Sheet refSheet = this.workbookPart.Workbook.Sheets.Elements<Spreadsheet.Sheet>()
                .LastOrDefault();

            return this.InsertWorksheetAfter(name, refSheet);
        }

        public IWorksheet InsertWorksheet(string name, int index)
        {
            Spreadsheet.Sheet refSheet = this.workbookPart.Workbook.Sheets.Elements<Spreadsheet.Sheet>()
                .Skip(index).LastOrDefault();

            return this.InsertWorksheetAfter(name, refSheet);
        }

        public IWorksheet PrependWorksheet(string name)
        {
            return this.InsertWorksheetAfter(name, null);
        }

        private IWorksheet InsertWorksheetAfter(string name, Spreadsheet.Sheet refSheet)
        {
            Packaging.WorksheetPart worksheetPart = this.workbookPart.AddNewPart<Packaging.WorksheetPart>();
            worksheetPart.Worksheet = new Spreadsheet.Worksheet(new Spreadsheet.SheetData());

            uint sheetId = (refSheet?.SheetId ?? 0u) + 1u;
            Spreadsheet.Sheet sheet = this.workbookPart.Workbook.Sheets.InsertAfter(
                new Spreadsheet.Sheet() { Id = this.workbookPart.GetIdOfPart(worksheetPart), SheetId = sheetId, Name = name },
                refSheet
            );

            sheet = sheet.NextSibling<Spreadsheet.Sheet>();
            ++sheetId;
            while (sheet != null)
            {
                sheet.SheetId = sheetId;

                sheet = sheet.NextSibling<Spreadsheet.Sheet>();
                ++sheetId;
            }

            return new Worksheet(worksheetPart);
        }
    }
}
