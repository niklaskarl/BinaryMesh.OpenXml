using System;
using DocumentFormat.OpenXml;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal interface IOpenXmlDataLabelAdjustElement
    {
        OpenXmlElement GetDataLabelAdjust();

        OpenXmlElement GetOrCreateDataLabelAdjust();
    }
}
