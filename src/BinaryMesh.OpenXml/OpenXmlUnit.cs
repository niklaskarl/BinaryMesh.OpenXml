using System;

namespace BinaryMesh.OpenXml
{
    public struct OpenXmlUnit
    {
        private readonly long emu;

        public OpenXmlUnit(long emu)
        {
            this.emu = emu;
        }

        public static OpenXmlUnit Cm(double cm)
        {
            return new OpenXmlUnit((long)(cm * 360000));
        }

        public static OpenXmlUnit Inch(double inch)
        {
            return new OpenXmlUnit((long)(inch * 914400));
        }

        public static OpenXmlUnit Points(double points)
        {
            return new OpenXmlUnit((long)(points * 12700));
        }

        public static implicit operator OpenXmlUnit(long emu)
        {
            return new OpenXmlUnit(emu);
        }

        public static implicit operator long(OpenXmlUnit unit)
        {
            return unit.emu;
        }

        public double AsCm()
        {
            return this.emu / 360000.0;
        }

        public double AsInch()
        {
            return this.emu / 914400.0;
        }

        public double AsPoints()
        {
            return this.emu / 12700.0;
        }

        public static OpenXmlUnit operator +(OpenXmlUnit left, OpenXmlUnit right)
        {
            return left.emu + right.emu;
        }

        public static OpenXmlUnit operator -(OpenXmlUnit left, OpenXmlUnit right)
        {
            return left.emu - right.emu;
        }

        public static OpenXmlUnit operator *(OpenXmlUnit left, OpenXmlUnit right)
        {
            return left.emu * right.emu;
        }

        public static OpenXmlUnit operator /(OpenXmlUnit left, OpenXmlUnit right)
        {
            return left.emu / right.emu;
        }
    }
}
