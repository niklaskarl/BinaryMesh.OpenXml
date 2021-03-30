using System;

namespace BinaryMesh.OpenXml
{
    public struct OpenXmlSize
    {
        private readonly OpenXmlUnit width;

        private readonly OpenXmlUnit height;

        public OpenXmlSize(OpenXmlUnit width, OpenXmlUnit height)
        {
            this.width = width;
            this.height = height;
        }

        public OpenXmlUnit Width => this.width;

        public OpenXmlUnit Height => this.height;

        public OpenXmlSize WithWidth(OpenXmlUnit width)
        {
            return new OpenXmlSize(
                width,
                this.height
            );
        }

        public OpenXmlSize WithHeight(OpenXmlUnit height)
        {
            return new OpenXmlSize(
                this.width,
                height
            );
        }
    }
}
