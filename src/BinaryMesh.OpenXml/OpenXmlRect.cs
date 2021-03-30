using System;

namespace BinaryMesh.OpenXml
{
    public struct OpenXmlRect
    {
        private readonly OpenXmlUnit left;

        private readonly OpenXmlUnit top;

        private readonly OpenXmlUnit width;

        private readonly OpenXmlUnit height;

        public OpenXmlRect(OpenXmlUnit left, OpenXmlUnit top, OpenXmlUnit width, OpenXmlUnit height)
        {
            this.left = left;
            this.top = top;
            this.width = width;
            this.height = height;
        }

        public OpenXmlUnit Left => this.left;

        public OpenXmlUnit Top => this.top;

        public OpenXmlUnit Right => this.left + this.width;

        public OpenXmlUnit Bottom => this.top + this.height;

        public OpenXmlUnit Width => this.width;

        public OpenXmlUnit Height => this.height;

        public OpenXmlRect WithLeft(OpenXmlUnit left)
        {
            return new OpenXmlRect(
                left,
                this.top,
                this.width,
                this.height
            );
        }

        public OpenXmlRect WithTop(OpenXmlUnit top)
        {
            return new OpenXmlRect(
                this.left,
                top,
                this.width,
                this.height
            );
        }

        public OpenXmlRect WithWidth(OpenXmlUnit width)
        {
            return new OpenXmlRect(
                this.left,
                this.top,
                width,
                this.height
            );
        }

        public OpenXmlRect WithHeight(OpenXmlUnit height)
        {
            return new OpenXmlRect(
                this.left,
                this.top,
                this.width,
                height
            );
        }

        public OpenXmlRect Translate(OpenXmlUnit x, OpenXmlUnit y)
        {
            return new OpenXmlRect(
                this.left + x,
                this.top + y,
                this.width,
                this.height
            );
        }

        public OpenXmlRect AddMargin(OpenXmlUnit size)
        {
            return new OpenXmlRect(
                this.left + size,
                this.top + size,
                this.width - size - size,
                this.height - size - size
            );
        }

        public OpenXmlRect AddMargin(OpenXmlUnit left, OpenXmlUnit top)
        {
            return new OpenXmlRect(
                this.left + left,
                this.top + top,
                this.width - left - left,
                this.height - top - top
            );
        }

        public OpenXmlRect AddMargin(OpenXmlUnit left, OpenXmlUnit top, OpenXmlUnit right, OpenXmlUnit bottom)
        {
            return new OpenXmlRect(
                this.left + left,
                this.top + top,
                this.width - left - right,
                this.height - top - bottom
            );
        }
    }
}
