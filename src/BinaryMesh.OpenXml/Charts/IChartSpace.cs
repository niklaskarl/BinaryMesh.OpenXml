using System;
using System.Collections.Generic;
using BinaryMesh.OpenXml.Spreadsheets;

namespace BinaryMesh.OpenXml.Charts
{
    public struct CartesianAxes
    {
        private readonly IChartAxis categoryAxis;

        private readonly IChartAxis valueAxis;

        public CartesianAxes(IChartAxis categoryAxis, IChartAxis valueAxis)
        {
            this.categoryAxis = categoryAxis;
            this.valueAxis = valueAxis;
        }

        public IChartAxis CategoryAxis => this.categoryAxis;

        public IChartAxis ValueAxis => this.valueAxis;
    }

    public interface IChartSpace
    {
        ISpreadsheetDocument OpenSpreadsheetDocument();

        IReadOnlyList<IChartAxis> CategoryAxes { get; }

        IReadOnlyList<IChartAxis> ValueAxes { get; }

        CartesianAxes AppendCartesianAxes();

        IPieChart InsertPieChart();

        IBarChart InsertBarChart(CartesianAxes axes);

        ILineChart InsertLineChart(CartesianAxes axes);
    }
}
