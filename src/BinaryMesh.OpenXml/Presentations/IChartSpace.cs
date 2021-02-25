using System;

using BinaryMesh.OpenXml.Spreadsheets;

namespace BinaryMesh.OpenXml.Presentations
{
    public interface IChartSpace
    {
        ISpreadsheetDocument OpenSpreadsheetDocument();

        IPieChart InsertPieChart();

        IBarChart InsertBarChart();

        IColumnChart InsertColumnChart();
    }
}
