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
        }
        
        internal override OpenXmlElement CreateColorElement()
        {
            return new RgbColorModelHex()
            {
                Val = $"{this.code:X}"
            };
        }
    }
}
