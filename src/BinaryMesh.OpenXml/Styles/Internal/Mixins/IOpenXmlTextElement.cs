using System;
using DocumentFormat.OpenXml;

namespace BinaryMesh.OpenXml.Styles.Internal.Mixins
{
    internal interface IOpenXmlTextElement
    {
        OpenXmlElement GetTextBody();

        OpenXmlElement GetOrCreateTextBody();
    }
}
