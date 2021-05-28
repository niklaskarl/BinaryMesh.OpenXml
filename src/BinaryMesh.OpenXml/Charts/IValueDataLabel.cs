using System;

using BinaryMesh.OpenXml.Styles;

namespace BinaryMesh.OpenXml.Charts
{

    public interface IValueDataLabel<out TFluent>
    {
        IVisualStyle<TFluent> Style { get; }

        ITextStyle<TFluent> Text { get; }

        TFluent SetDelete(bool show);

        TFluent Clear();
    }
}
