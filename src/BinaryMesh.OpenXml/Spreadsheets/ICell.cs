using System;

namespace BinaryMesh.OpenXml.Spreadsheets
{
    public interface ICell
    {
        IWorksheet Worksheet { get; }

        uint Column { get; }

        bool IsColumnFixed { get; }

        uint Row { get; }

        bool IsRowFixed { get; }

        string Reference { get; }

        String InnerValue { get; }

        ICell SetValue(double value);

        ICell SetValue(string value);
    }
}
