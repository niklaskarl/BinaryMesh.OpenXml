using System;
using System.Text;
using System.Text.RegularExpressions;

namespace BinaryMesh.OpenXml.Spreadsheets.Helpers
{
    internal static class ReferenceEncoder
    {

        public static string EncodeRangeReference(
            string worksheetName,
            uint? startColumn, bool isStartColumnFixed,
            uint? startRow, bool isStartRowFixed,
            uint? endColumn, bool isEndColumnFixed,
            uint? endRow, bool isEndRowFixed)
        {
            string localRangeReference = EncodeLocalRangeReference(
                startColumn, isStartColumnFixed,
                startRow, isStartRowFixed,
                endColumn, isEndColumnFixed,
                endRow, isEndRowFixed
            );

            return worksheetName != null ? $"{worksheetName}!{localRangeReference}" : localRangeReference;
        }

        public static bool TryDecodeRangeReference(
            string reference,
            out string worksheet,
            out uint? startColumn, out bool isStartColumnFixed,
            out uint? startRow, out bool isStartRowFixed,
            out uint? endColumn, out bool isEndColumnFixed,
            out uint? endRow, out bool isEndRowFixed)
        {
            int index = reference.IndexOf('!');
            worksheet = index >= 0 ? reference.Remove(index) : null;
            reference = index >= 0 ? reference.Substring(index + 1) : reference;

            return TryDecodeLocalRangeReference(
                reference,
                out startColumn, out isStartColumnFixed,
                out startRow, out isStartRowFixed,
                out endColumn, out isEndColumnFixed,
                out endRow, out isEndRowFixed
            );
        }

        public static string EncodeLocalRangeReference(
            uint? startColumn, bool isStartColumnFixed,
            uint? startRow, bool isStartRowFixed,
            uint? endColumn, bool isEndColumnFixed,
            uint? endRow, bool isEndRowFixed)
        {
            StringBuilder builder = new StringBuilder();
            if (startColumn.HasValue)
                {
                    if (isStartColumnFixed)
                    {
                        builder.Append('$');
                    }

                    builder.Append(ReferenceEncoder.EncodeColumnReference(startColumn.Value));
                }

                if (startRow.HasValue)
                {
                    if (isStartRowFixed)
                    {
                        builder.Append('$');
                    }

                    builder.Append(startRow.Value + 1);
                }

                if (endColumn != startColumn || endRow != startRow)
                {
                    builder.Append(':');

                    if (endColumn.HasValue)
                    {
                        if (isEndColumnFixed)
                        {
                            builder.Append('$');
                        }

                    builder.Append(ReferenceEncoder.EncodeColumnReference(endColumn.Value));
                    }

                    if (endRow.HasValue)
                    {
                        if (isEndRowFixed)
                        {
                            builder.Append('$');
                        }

                        builder.Append(endRow.Value + 1);
                    }
                }

                return builder.ToString();
        }

        public static bool TryDecodeLocalRangeReference(
            string reference,
            out uint? startColumn, out bool isStartColumnFixed,
            out uint? startRow, out bool isStartRowFixed,
            out uint? endColumn, out bool isEndColumnFixed,
            out uint? endRow, out bool isEndRowFixed)
        {
            bool result = false;

            int index = reference.IndexOf(':');
            if (index >= 0)
            {
                string start = reference.Remove(index);
                string end = reference.Substring(index + 1);
                bool startResult = TryDecodePartialRangeReference(start, out startColumn, out isStartColumnFixed, out startRow, out isStartRowFixed);
                bool endResult = TryDecodePartialRangeReference(end, out endColumn, out isEndColumnFixed, out endRow, out isEndRowFixed);
                result = startResult && endResult;
            }
            else
            {
                result = TryDecodePartialRangeReference(reference, out uint? column, out bool isColumnFixed, out uint? row, out bool isRowFixed);
                startColumn = column;
                isStartColumnFixed = isColumnFixed;
                startRow = row;
                isStartRowFixed = isRowFixed;
                endColumn = column;
                isEndColumnFixed = isColumnFixed;
                endRow = row;
                isEndRowFixed = isRowFixed;
                result = true;
            }

            return result;
        }

        private static bool TryDecodePartialRangeReference(
            string reference,
            out uint? column, out bool isColumnFixed,
            out uint? row, out bool isRowFixed)
        {
            column = null;
            isColumnFixed = false;
            row = null;
            isRowFixed = false;

            Regex pattern = new Regex("^((?<isColumnFixed>\\$?)(?<column>[a-zA-Z]+))?((?<isRowFixed>\\$?)(?<row>[0-9]+))?$");
            Match match = pattern.Match(reference);
            if (match.Success)
            {
                if (match.Groups["column"].Success)
                {
                    isColumnFixed = match.Groups["isColumnFixed"].Success;
                    column = DecodeColumnReference(match.Groups["column"].Value);
                }

                if (match.Groups["row"].Success)
                {
                    isRowFixed = match.Groups["isRowFixed"].Success;
                    row = uint.Parse(match.Groups["row"].Value) - 1;
                }

                return true;
            }

            return false;
        }

        public static bool TryDecodeCellReference(
            string reference,
            out uint column, out bool isColumnFixed,
            out uint row, out bool isRowFixed)
        {
            bool result = TryDecodePartialRangeReference(reference, out uint? optColumn, out isColumnFixed, out uint? optRow, out isRowFixed);
            column = optColumn ?? 0u;
            row = optRow ?? 0u;

            return result && optColumn.HasValue && optRow.HasValue;
        }

        public static string EncodeColumnReference(uint column)
        {
            string result = ((char)('A' + (column) % ('Z' - 'A' + 1))).ToString();
            while((column = (column) / ('Z' - 'A' + 1)) > 0)
            {
                result = ((char)('A' + (column) % ('Z' - 'A' + 1))) + result;
            }

            return result;
        }

        public static uint DecodeColumnReference(string reference)
        {
            uint column = 0;

            int i = 0;
            while (i < reference.Length && ((reference[i] >= 'A' && reference[i] <= 'Z') || (reference[i] >= 'a' && reference[i] <= 'z')))
            {
                if (reference[i] < 'a')
                {
                    column *= 'Z' - 'A' + 1;
                    column += ((uint)(reference[i] - 'A') + 1u);
                }
                else
                {
                    column *= 'z' - 'a' + 1;
                    column += ((uint)(reference[i] - 'a') + 1u);
                }

                ++i;
            }

            if (i != reference.Length)
            {
                throw new FormatException();
            }

            return column - 1;
        }
    }
}
