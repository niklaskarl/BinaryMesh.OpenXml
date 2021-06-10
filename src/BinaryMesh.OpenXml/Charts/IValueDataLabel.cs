using System;

using BinaryMesh.OpenXml.Styles;

namespace BinaryMesh.OpenXml.Charts
{

    public interface IValueDataLabel<out TFluent> : IDataLabel<TFluent>
    {
        TFluent SetDelete(bool show);
    }
}
