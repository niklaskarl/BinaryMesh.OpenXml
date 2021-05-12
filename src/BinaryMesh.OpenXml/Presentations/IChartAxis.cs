using System;

using BinaryMesh.OpenXml.Spreadsheets;

namespace BinaryMesh.OpenXml.Presentations
{
    public interface IChartAxis
    {
        uint Id { get; }

        IChartAxis SetVisibility(bool value);

        IVisualStyle<IChartAxis> Style { get; }

        ITextStyle<IChartAxis> Text { get; }
    }
}
