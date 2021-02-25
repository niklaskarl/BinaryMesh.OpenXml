using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using Packaging = DocumentFormat.OpenXml.Packaging;

using BinaryMesh.OpenXml.Helpers;

namespace BinaryMesh.OpenXml.Spreadsheets.Internal
{
    internal class SpreadsheetDocument : ISpreadsheetDocument, IDisposable
    {
        private readonly Stream stream;

        private readonly bool keepStreamOpen;

        private readonly Packaging.SpreadsheetDocument spreadsheetDocument;

        public SpreadsheetDocument()
        {
            this.stream = new MemoryStream();
            this.keepStreamOpen = false;

            this.spreadsheetDocument = Packaging.SpreadsheetDocument.Create(this.stream, SpreadsheetDocumentType.Workbook);
        }

        public SpreadsheetDocument(Stream stream)
        {
            this.stream = stream;
            this.keepStreamOpen = false;

            this.spreadsheetDocument = Packaging.SpreadsheetDocument.Open(this.stream, true);
        }

        public SpreadsheetDocument(Stream stream, bool keepStreamOpen)
        {
            this.stream = stream;
            this.keepStreamOpen = keepStreamOpen;

            this.spreadsheetDocument = Packaging.SpreadsheetDocument.Open(this.stream, true);
        }

        public IWorkbook Workbook => new Workbook(this.spreadsheetDocument.WorkbookPart ?? this.spreadsheetDocument.AddWorkbookPart());

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
