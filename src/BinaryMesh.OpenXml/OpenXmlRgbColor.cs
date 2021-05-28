using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;

namespace BinaryMesh.OpenXml
{
    internal sealed class OpenXmlRgbColor : OpenXmlColor
    {
        private uint code;

        internal OpenXmlRgbColor(uint code)
        {
            this.code = code;
            this.alpha = 1.0;
        }

        internal OpenXmlRgbColor(OpenXmlRgbColor other)
            : base (other)
        {
            this.code = other.code;
        }

        internal override OpenXmlElement CreateColorElement()
        {
            RgbColorModelHex element = new RgbColorModelHex()
            {
                Val = $"{this.code:X6}"
            };

            this.AnnotateOpenXmlElement(element);
            
            return element;
        }

        internal override OpenXmlColor Clone()
        {
            return new OpenXmlRgbColor(this);
        }
    }
}
