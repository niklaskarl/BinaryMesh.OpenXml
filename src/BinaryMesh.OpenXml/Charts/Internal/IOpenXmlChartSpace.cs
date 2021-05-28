using System;
using Packaging = DocumentFormat.OpenXml.Packaging;

namespace BinaryMesh.OpenXml.Charts.Internal
{
    internal interface IOpenXmlChartSpace : IChartSpace
    {
        Packaging.ChartPart ChartPart { get; }
    }
}
