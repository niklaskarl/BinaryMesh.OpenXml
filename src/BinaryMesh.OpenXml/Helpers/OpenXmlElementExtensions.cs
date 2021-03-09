using System;
using System.Collections.Generic;
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

        public static TElement AppendFluent<TElement>(this TElement element, IEnumerable<OpenXmlElement> children) where TElement : OpenXmlElement
        {
            element.Append(children);
            return element;
        }
    }
}
