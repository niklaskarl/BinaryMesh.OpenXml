using System;
using System.IO;

namespace BinaryMesh.OpenXml.Spreadsheets
{
    public interface ISpreadsheetDocument : IDisposable
    {
        IWorkbook Workbook { get; }

        void Close(Stream destination);
    }
}
