using System;

namespace BinaryMesh.OpenXml.Spreadsheets
{
    public interface ICell
    {
        uint Column { get; }

        bool IsColumnFixed { get; }

        uint Row { get; }

        bool IsRowFixed { get; }

        string Reference { get; }

        ICell SetValue(double value);

        ICell SetValue(string value);
    }
}
