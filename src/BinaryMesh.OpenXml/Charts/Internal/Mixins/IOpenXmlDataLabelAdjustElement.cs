using System;
using DocumentFormat.OpenXml;

namespace BinaryMesh.OpenXml.Charts.Internal.Mixins
{
    internal interface IOpenXmlDataLabelAdjustElement
    {
        OpenXmlElement GetDataLabelAdjust();

        OpenXmlElement GetOrCreateDataLabelAdjust();
    }
}
