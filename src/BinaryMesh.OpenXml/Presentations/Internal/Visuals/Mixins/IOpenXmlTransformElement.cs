using System;
using DocumentFormat.OpenXml;

namespace BinaryMesh.OpenXml.Presentations.Internal.Mixins
{
    internal interface IOpenXmlTransformElement
    {
        OpenXmlElement GetTransform();

        OpenXmlElement GetOrCreateTransform();
    }
}
