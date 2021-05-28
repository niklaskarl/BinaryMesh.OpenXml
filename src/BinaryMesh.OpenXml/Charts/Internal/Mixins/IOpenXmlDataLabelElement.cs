using System;
using DocumentFormat.OpenXml;

namespace BinaryMesh.OpenXml.Charts.Internal.Mixins
{
    internal interface IOpenXmlDataLabelElement
    {
        OpenXmlElement GetDataLabel();

        OpenXmlElement GetOrCreateDataLabel();
    }
}
