using System;
using DocumentFormat.OpenXml;

namespace BinaryMesh.OpenXml
{
    internal static class OpenXmlElementExtensions
    {
        public static TElement AppendChildFluent<TElement>(this TElement element, OpenXmlElement child) where TElement : OpenXmlElement
        {
            element.AppendChild(child);
            return element;
        }
    }
}
