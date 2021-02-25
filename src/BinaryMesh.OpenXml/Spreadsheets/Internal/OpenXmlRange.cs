using System;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Spreadsheet;

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

        public ICell this[uint column, uint row] => throw new NotImplementedException();

        public string Formula
        {
            get
            {
                return null;
            }
        }

        public int? Width => throw new NotImplementedException();

        public int? Height => throw new NotImplementedException();
    }
}
