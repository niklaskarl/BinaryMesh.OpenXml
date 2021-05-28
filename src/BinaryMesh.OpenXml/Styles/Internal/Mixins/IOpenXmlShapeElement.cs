using System;
using DocumentFormat.OpenXml;

namespace BinaryMesh.OpenXml.Styles.Internal.Mixins
{
    internal interface IOpenXmlShapeElement
    {
        OpenXmlElement GetShapeProperties();

        OpenXmlElement GetOrCreateShapeProperties();
    }
}
