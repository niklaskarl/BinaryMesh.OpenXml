using System;

namespace BinaryMesh.OpenXml.Spreadsheets
{
    public interface IRange
    {
        ICell this[uint column, uint row] { get; }

        IWorksheet Worksheet { get; }

        string Formula { get; }

        uint? StartColumn { get; }

        uint? StartRow { get; }

        uint? EndColumn { get; }

        uint? EndRow { get; }

        int? Width { get; }

        int? Height { get; }
    }
}
