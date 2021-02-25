using System;

using BinaryMesh.OpenXml.Spreadsheets;

namespace BinaryMesh.OpenXml.Presentations
{
    public interface IChartSeries
    {
        IChartSeries SetText(IRange range);

        IChartSeries SetCategoryAxis(IRange range);

        IChartSeries SetValueAxis(IRange range);
    }
}
