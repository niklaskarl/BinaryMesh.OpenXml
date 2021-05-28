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

        internal OpenXmlSchemeColor(OpenXmlSchemeColor other)
            : base (other)
        {
            this.color = other.color;
        }
        
        internal override OpenXmlElement CreateColorElement()
        {
            SchemeColor element = new SchemeColor()
            {
                Val = color
            };

            this.AnnotateOpenXmlElement(element);

            return element;
        }

        internal override OpenXmlColor Clone()
        {
            return new OpenXmlSchemeColor(this);
        }
    }
}
