using System;
using System.Text.RegularExpressions;

namespace BinaryMesh.OpenXml.Spreadsheets.Helpers
{
    internal static class ReferenceDecoder
    {
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
                    column = DecodeColumn(match.Groups["column"].Value);
                }

                if (match.Groups["row"].Success)
                {
                    isRowFixed = match.Groups["isRowFixed"].Success;
                    row = uint.Parse(match.Groups["row"].Value);
                }

                return true;
            }

            return false;
        }

        private static uint DecodeColumn(string reference)
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

            return column;
        }
    }
}
