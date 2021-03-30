using System;

namespace BinaryMesh.OpenXml
{
    public struct OpenXmlPoint
    {
        private readonly OpenXmlUnit left;

        private readonly OpenXmlUnit top;

        public OpenXmlPoint(OpenXmlUnit left, OpenXmlUnit top)
        {
            this.left = left;
            this.top = top;
        }

        public OpenXmlUnit Left => this.left;

        public OpenXmlUnit Top => this.top;

        public OpenXmlPoint WithLeft(OpenXmlUnit left)
        {
            return new OpenXmlPoint(
                left,
                this.top
            );
        }

        public OpenXmlPoint WithTop(OpenXmlUnit top)
        {
            return new OpenXmlPoint(
                this.left,
                top
            );
        }

        public OpenXmlPoint Translate(OpenXmlUnit x, OpenXmlUnit y)
        {
            return new OpenXmlPoint(
                this.left + x,
                this.top + y
            );
        }
    }
}
