using System;

namespace BinaryMesh.OpenXml.Internal
{
    internal interface IOpenXmlDocument
    {
        IOpenXmlTextStyle DefaultTextStyle { get; }

        IOpenXmlTheme Theme { get; }
    }
}
