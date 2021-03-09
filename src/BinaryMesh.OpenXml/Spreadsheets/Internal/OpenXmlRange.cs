using System;
using System.Text;

using BinaryMesh.OpenXml.Spreadsheets.Helpers;

namespace BinaryMesh.OpenXml.Spreadsheets.Internal
{
    internal sealed class OpenXmlRange : IRange
    {
        private readonly OpenXmlWorksheet worksheet;

        private readonly uint? startColumn;

        private readonly bool isStartColumnFixed;

        private readonly uint? startRow;

        private readonly bool isStartRowFixed;

        private readonly uint? endColumn;

        private readonly bool isEndColumnFixed;

        private readonly uint? endRow;

        private readonly bool isEndRowFixed;

        public OpenXmlRange(OpenXmlWorksheet worksheet, uint? startColumn, uint? startRow, uint? endColumn, uint? endRow)
        {
            this.worksheet = worksheet;
            this.startColumn = startColumn;
            this.isStartColumnFixed = false;
            this.startRow = startRow;
            this.isStartRowFixed = false;
            this.endColumn = endColumn;
            this.isEndColumnFixed = false;
            this.endRow = endRow;
            this.isEndRowFixed = false;
        }

        public OpenXmlRange(
            OpenXmlWorksheet worksheet,
            uint? startColumn, bool isStartColumnFixed,
            uint? startRow, bool isStartRowFixed,
            uint? endColumn, bool isEndColumnFixed,
            uint? endRow, bool isEndRowFixed)
        {
            this.worksheet = worksheet;
            this.startColumn = startColumn;
            this.isStartColumnFixed = isStartColumnFixed;
            this.startRow = startRow;
            this.isStartRowFixed = isStartRowFixed;
            this.endColumn = endColumn;
            this.isEndColumnFixed = isEndColumnFixed;
            this.endRow = endRow;
            this.isEndRowFixed = isEndRowFixed;
        }

        public ICell this[uint column, uint row] =>
            (!this.Width.HasValue || column < this.Width.Value) && (!this.Height.HasValue || row < this.Height.Value) ?
                this.worksheet.Cells[(this.startColumn ?? 0u) + column, (this.startRow ?? 0u) + row] : throw new ArgumentOutOfRangeException();

        public string Formula => ReferenceEncoder.EncodeRangeReference(this.worksheet?.Name, this.startColumn, this.isStartColumnFixed, this.startRow, this.isStartRowFixed, this.endColumn, this.isEndColumnFixed, this.endRow, this.isEndRowFixed);

        public IWorksheet Worksheet => this.worksheet;

        public uint? StartColumn => this.startColumn;

        public uint? StartRow => this.startRow;

        public uint? EndColumn => this.endColumn;

        public uint? EndRow => this.endRow;

        public int? Width => (this.startColumn.HasValue && this.endColumn.HasValue) ? ((int)this.endColumn.Value - (int)this.startColumn.Value + 1) : (int?)null;

        public int? Height => (this.startRow.HasValue && this.endRow.HasValue) ? ((int)this.endRow.Value - (int)this.startRow.Value + 1) : (int?)null;
    }
}
