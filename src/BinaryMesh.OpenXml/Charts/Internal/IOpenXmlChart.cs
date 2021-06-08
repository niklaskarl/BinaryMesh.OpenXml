using System;

namespace BinaryMesh.OpenXml.Charts.Internal
{
    internal interface IOpenXmlChart : IChart
    {
        uint SeriesCount { get; }
    }
}
