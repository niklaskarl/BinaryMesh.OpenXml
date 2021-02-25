using System;

namespace BinaryMesh.OpenXml.Spreadsheets
{
    public interface IWorksheetCells
    {
        ICell this[uint column, uint row] { get; }

        ICell this[string reference] { get; }
    }
}
