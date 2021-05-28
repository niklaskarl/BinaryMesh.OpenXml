using System;

using BinaryMesh.OpenXml.Styles;

namespace BinaryMesh.OpenXml.Charts
{
    public interface IDataLabel<out TFluent>
    {
        IVisualStyle<TFluent> Style { get; }

        ITextStyle<TFluent> Text { get; }

        TFluent SetShowValue(bool show);

        TFluent SetShowPercent(bool show);

        TFluent SetShowCategoryName(bool show);

        TFluent SetShowLegendKey(bool show);

        TFluent SetShowSeriesName(bool show);

        TFluent SetShowBubbleSize(bool show);

        TFluent SetShowLeaderLines(bool show);

        TFluent Clear();
    }
}
