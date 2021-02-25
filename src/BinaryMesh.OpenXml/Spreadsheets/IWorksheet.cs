using System;

namespace BinaryMesh.OpenXml.Spreadsheets
{
    public interface IWorksheet
    {
        IWorksheetCells Cells { get; }
    }
}
