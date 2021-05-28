using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;

namespace BinaryMesh.OpenXml
{
    public abstract class OpenXmlColor
    {
        protected double luminanceModulation;

        protected double luminanceOffset;

        protected double alpha;

        internal OpenXmlColor()
        {
            this.luminanceModulation = 1.0;
            this.luminanceOffset = 0.0;
            this.alpha = 1.0;
        }

        internal OpenXmlColor(OpenXmlColor other)
        {
            this.alpha = other.alpha;
        }

        public double LuminanceModulation => luminanceModulation;

        public double LuminanceOffset => luminanceOffset;

        public double Alpha => this.alpha;

        internal abstract OpenXmlColor Clone();

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

        public OpenXmlColor WithLuminanceModulation(double luminanceModulation)
        {
            OpenXmlColor color = this.Clone();
            color.luminanceModulation = luminanceModulation;

            return color;
        }

        public OpenXmlColor WithLuminanceOffset(double luminanceOffset)
        {
            OpenXmlColor color = this.Clone();
            color.luminanceOffset = luminanceOffset;

            return color;
        }

        public OpenXmlColor WithAlpha(double alpha)
        {
            OpenXmlColor color = this.Clone();
            color.alpha = alpha;

            return color;
        }

        protected void AnnotateOpenXmlElement(OpenXmlElement element)
        {
            if (luminanceModulation != 1.0)
            {
                element.AppendChild(new LuminanceModulation() { Val = (int)(luminanceModulation * 100000) });
            }

            if (luminanceOffset != 0.0)
            {
                element.AppendChild(new LuminanceOffset() { Val = (int)(luminanceOffset * 100000) });
            }

            if (alpha != 1.0)
            {
                element.AppendChild(new Alpha() { Val = (int)(alpha * 100000) });
            }
        }
    }
}
