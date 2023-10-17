using System;

namespace BinaryMesh.OpenXml.Spreadsheets
{
    public interface IRowIterator : IDisposable
    {
        uint CurrentRowIndex { get; }

        uint CurrentColumnIndex { get; }

        string CurrentReference { get; }

        void NextRow();

        void NextColumn();

        void SetValue(double value);

        void SetValue(string value);
    }
}
