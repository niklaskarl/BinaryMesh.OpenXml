using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;

namespace BinaryMesh.OpenXml
{
    public abstract class OpenXmlColor
    {
        internal OpenXmlColor()
        {
        }

        internal abstract OpenXmlElement CreateColorElement();

        public static OpenXmlColor Light1 => new OpenXmlSchemeColor(SchemeColorValues.Light1);

        public static OpenXmlColor Light2 => new OpenXmlSchemeColor(SchemeColorValues.Light2);

        public static OpenXmlColor Dark1 => new OpenXmlSchemeColor(SchemeColorValues.Dark1);

        public static OpenXmlColor Dark2 => new OpenXmlSchemeColor(SchemeColorValues.Dark2);

        public static OpenXmlColor Text1 => new OpenXmlSchemeColor(SchemeColorValues.Text1);

        public static OpenXmlColor Text2 => new OpenXmlSchemeColor(SchemeColorValues.Text2);

        public static OpenXmlColor Accent1 => new OpenXmlSchemeColor(SchemeColorValues.Accent1);

        public static OpenXmlColor Accent2 => new OpenXmlSchemeColor(SchemeColorValues.Accent2);

        public static OpenXmlColor Accent3 => new OpenXmlSchemeColor(SchemeColorValues.Accent3);

        public static OpenXmlColor Accent4 => new OpenXmlSchemeColor(SchemeColorValues.Accent4);

        public static OpenXmlColor Accent5 => new OpenXmlSchemeColor(SchemeColorValues.Accent5);

        public static OpenXmlColor Accent6 => new OpenXmlSchemeColor(SchemeColorValues.Accent6);

        public static OpenXmlColor Rgb(uint code)
        {
            return new OpenXmlRgbColor(code);
        }

        public static OpenXmlColor Rgb(byte red, byte green, byte blue)
        {
            return new OpenXmlRgbColor(((uint)red << 16) | ((uint)green << 8) | ((uint)blue << 0));
        }
    }
}
