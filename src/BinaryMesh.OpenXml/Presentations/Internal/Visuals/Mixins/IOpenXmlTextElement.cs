using System;
using DocumentFormat.OpenXml;

namespace BinaryMesh.OpenXml.Presentations.Internal.Mixins
{
    internal interface IOpenXmlTextElement
    {
        OpenXmlElement GetTextBody();

        OpenXmlElement GetOrCreateTextBody();
    }
}
