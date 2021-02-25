using System;
using System.Collections.Generic;

namespace BinaryMesh.OpenXml.Spreadsheets
{
    public interface IWorkbook
    {
        KeyedReadOnlyList<string, IWorksheet> Worksheets { get; }

        IWorksheet PrependWorksheet(string name);

        IWorksheet InsertWorksheet(string name, int index);

        IWorksheet AppendWorksheet(string name);

        IRange GetRange(string formula);
    }
}
