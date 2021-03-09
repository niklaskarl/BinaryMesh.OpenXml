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

        public static implicit operator OpenXmlUnit(long emu)
        {
            return new OpenXmlUnit(emu);
        }

        public static implicit operator long(OpenXmlUnit unit)
        {
            return unit.emu;
        }
    }
}
