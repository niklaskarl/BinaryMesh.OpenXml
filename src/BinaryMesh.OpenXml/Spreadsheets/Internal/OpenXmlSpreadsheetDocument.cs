using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace BinaryMesh.OpenXml.Spreadsheets.Internal
{
    internal class OpenXmlSpreadsheetDocument : ISpreadsheetDocument, IDisposable
    {
        private readonly Stream stream;

        private readonly bool keepStreamOpen;

        private readonly SpreadsheetDocument spreadsheetDocument;

        public OpenXmlSpreadsheetDocument()
        {
            this.stream = new MemoryStream();
            this.keepStreamOpen = false;

            this.spreadsheetDocument = SpreadsheetDocument.Create(this.stream, SpreadsheetDocumentType.Workbook);
        }

        public OpenXmlSpreadsheetDocument(Stream stream, bool initialize)
        {
            this.stream = stream;
            this.keepStreamOpen = false;

            this.spreadsheetDocument = initialize ?
                SpreadsheetDocument.Create(this.stream, SpreadsheetDocumentType.Workbook) :
                SpreadsheetDocument.Open(this.stream, true);
        }

        public OpenXmlSpreadsheetDocument(Stream stream, bool initialize, bool keepStreamOpen)
        {
            this.stream = stream;
            this.keepStreamOpen = keepStreamOpen;

            this.spreadsheetDocument = initialize ?
                SpreadsheetDocument.Create(this.stream, SpreadsheetDocumentType.Workbook) :
                SpreadsheetDocument.Open(this.stream, true);
        }

        public IWorkbook Workbook => new OpenXmlWorkbook(this.spreadsheetDocument.WorkbookPart ?? this.spreadsheetDocument.AddWorkbookPart());

        public void Close(Stream destination)
        {
            this.spreadsheetDocument.Close();
            this.stream.Position = 0;
            this.stream.CopyTo(destination);
            this.Dispose();
        }

        public void Dispose()
        {
            this.spreadsheetDocument.Dispose();
            if (!this.keepStreamOpen)
            {
                this.stream.Dispose();
            }
        }
    }
}
