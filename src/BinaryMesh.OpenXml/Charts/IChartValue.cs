using System;
using System.Collections.Generic;

using BinaryMesh.OpenXml.Styles;

namespace BinaryMesh.OpenXml.Charts
{
    public interface IChartValue<out TFluent>
    {
        IVisualStyle<TFluent> Style { get; }

        IValueDataLabel<TFluent> DataLabel { get; }
    }
}
