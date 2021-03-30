using System;
using DocumentFormat.OpenXml;

namespace BinaryMesh.OpenXml.Presentations.Internal.Mixins
{
    internal interface IOpenXmlShapeElement
    {
        OpenXmlElement GetShapeProperties();

        OpenXmlElement GetOrCreateShapeProperties();
    }
}
