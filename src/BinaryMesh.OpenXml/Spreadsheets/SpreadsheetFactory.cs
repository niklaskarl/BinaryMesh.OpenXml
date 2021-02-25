using System;
using System.IO;
using BinaryMesh.OpenXml.Spreadsheets.Internal;

namespace BinaryMesh.OpenXml.Spreadsheets
{
    public static class SpreadsheetFactory
    {
        public static ISpreadsheetDocument CreateSpreadsheet()
        {
            return new OpenXmlSpreadsheetDocument();
        }

        public static ISpreadsheetDocument OpenSpreadsheet(string source)
        {
            MemoryStream stream = new MemoryStream();
            using (Stream sourceStream = new FileStream(source, FileMode.Open, FileAccess.Read))
            {
                sourceStream.CopyTo(stream);
                stream.Seek(0, SeekOrigin.Begin);
            }
            
            return new OpenXmlSpreadsheetDocument(stream);
        }

        public static ISpreadsheetDocument OpenSpreadsheet(Stream sourceStream)
        {
            MemoryStream stream = new MemoryStream();
            sourceStream.CopyTo(stream);
            stream.Seek(0, SeekOrigin.Begin);
            
            return new OpenXmlSpreadsheetDocument(stream);
        }
    }
}
