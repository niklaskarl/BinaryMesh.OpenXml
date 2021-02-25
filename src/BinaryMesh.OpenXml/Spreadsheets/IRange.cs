using System;

namespace BinaryMesh.OpenXml.Spreadsheets
{
    public interface IRange
    {
        ICell this[uint column, uint row] { get; }

        string Formula { get; }

        int? Width { get; }

        int? Height { get; }
    }
}
