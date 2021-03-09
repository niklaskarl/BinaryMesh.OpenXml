using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;

namespace BinaryMesh.OpenXml
{
    internal sealed class OpenXmlSchemeColor : OpenXmlColor
    {
        private SchemeColorValues color;

        internal OpenXmlSchemeColor(SchemeColorValues color)
        {
            this.color = color;
        }
        
        internal override OpenXmlElement CreateColorElement()
        {
            return new SchemeColor()
            {
                Val = color
            };
        }
    }
}
