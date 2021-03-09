using System;

namespace BinaryMesh.OpenXml.Spreadsheets
{
    public interface IWorksheet
    {
        string Name { get; }

        IWorksheetCells Cells { get; }

        IRange GetRange(string formula);

        IRange GetRange(uint startColumn, uint startRow, uint endColumn, uint endRow);
    }
}
