using System;
using DocumentFormat.OpenXml;

namespace BinaryMesh.OpenXml.Styles.Internal.Mixins
{
    internal interface IOpenXmlTransformElement
    {
        OpenXmlElement GetTransform();

        OpenXmlElement GetOrCreateTransform();
    }
}
