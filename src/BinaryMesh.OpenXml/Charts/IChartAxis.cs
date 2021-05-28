using System;

using BinaryMesh.OpenXml.Styles;

namespace BinaryMesh.OpenXml.Charts
{
    public interface IChartAxisMajorGridlines<out TResult>
    {
        bool Exists { get; }

        IStrokeStyle<TResult> Style { get; }

        TResult Add();

        TResult Remove();
    }

    public interface IChartAxis
    {
        IChartAxis SetVisibility(bool value);

        uint Id { get; }

        IChartAxisMajorGridlines<IChartAxis> MajorGridlines { get; }

        IVisualStyle<IChartAxis> Style { get; }

        ITextStyle<IChartAxis> Text { get; }
    }
}
