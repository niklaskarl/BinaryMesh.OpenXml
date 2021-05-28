using System;
using System.Collections.Generic;

using BinaryMesh.OpenXml.Styles;

namespace BinaryMesh.OpenXml.Charts
{
    public interface IChartSeries<out TSeriesFluent, out TValueFluent>
        where TValueFluent : IChartValue<TValueFluent>
    {
        IReadOnlyList<IChartValue<TValueFluent>> Values { get; }

        IVisualStyle<TSeriesFluent> Style { get; }

        IDataLabel<TSeriesFluent> DataLabel { get; }
    }
}
