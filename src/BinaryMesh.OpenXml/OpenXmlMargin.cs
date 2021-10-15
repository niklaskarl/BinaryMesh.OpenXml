using System;

namespace BinaryMesh.OpenXml
{
    public struct OpenXmlMargin
    {
        private readonly OpenXmlUnit left;

        private readonly OpenXmlUnit top;

        private readonly OpenXmlUnit right;

        private readonly OpenXmlUnit bottom;

        public OpenXmlMargin(OpenXmlUnit left, OpenXmlUnit top, OpenXmlUnit right, OpenXmlUnit bottom)
        {
            this.left = left;
            this.top = top;
            this.right = right;
            this.bottom = bottom;
        }

        public OpenXmlUnit Left => this.left;

        public OpenXmlUnit Top => this.top;

        public OpenXmlUnit Right => this.right;

        public OpenXmlUnit Bottom => this.bottom;

        public OpenXmlMargin WithLeft(OpenXmlUnit left)
        {
            return new OpenXmlMargin(
                left,
                this.top,
                this.right,
                this.bottom
            );
        }

        public OpenXmlMargin WithTop(OpenXmlUnit top)
        {
            return new OpenXmlMargin(
                this.left,
                top,
                this.right,
                this.bottom
            );
        }

        public OpenXmlMargin WithRight(OpenXmlUnit right)
        {
            return new OpenXmlMargin(
                this.left,
                this.top,
                right,
                this.bottom
            );
        }

        public OpenXmlMargin WithBottom(OpenXmlUnit bottom)
        {
            return new OpenXmlMargin(
                this.left,
                this.top,
                this.right,
                bottom
            );
        }
    }
}
