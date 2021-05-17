using System;
using System.Collections.Generic;

namespace BinaryMesh.OpenXml.Presentations
{

    public interface IDataLabelAdjust<out TFluent>
    {
        IVisualStyle<TFluent> Style { get; }

        ITextStyle<TFluent> Text { get; }

        TFluent SetDelete(bool show);

        TFluent Clear();
    }

    public interface IChartValue<out TFluent>
    {
        IVisualStyle<TFluent> Style { get; }

        IDataLabelAdjust<TFluent> DataLabel { get; }
    }

    public interface IChartSeries<out TSeriesFluent, out TValueFluent>
        where TValueFluent : IChartValue<TValueFluent>
    {
        IReadOnlyList<IChartValue<TValueFluent>> Values { get; }

        IVisualStyle<TSeriesFluent> Style { get; }

        IDataLabel<TSeriesFluent> DataLabel { get; }
    }
}
