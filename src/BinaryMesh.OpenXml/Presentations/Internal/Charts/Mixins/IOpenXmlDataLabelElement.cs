using System;
using DocumentFormat.OpenXml;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal interface IOpenXmlDataLabelElement
    {
        OpenXmlElement GetDataLabel();

        OpenXmlElement GetOrCreateDataLabel();
    }
}
