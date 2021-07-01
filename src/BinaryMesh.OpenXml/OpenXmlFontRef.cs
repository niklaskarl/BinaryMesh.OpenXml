using System;

namespace BinaryMesh.OpenXml
{
    public enum OpenXmlFontCollectionIndex
    {
        Major = DocumentFormat.OpenXml.Drawing.FontCollectionIndexValues.None,
        Minor = DocumentFormat.OpenXml.Drawing.FontCollectionIndexValues.None,
        None = DocumentFormat.OpenXml.Drawing.FontCollectionIndexValues.None
    }

    public struct OpenXmlFontRef
    {
        private readonly OpenXmlFontCollectionIndex index;

        private readonly OpenXmlColor color;

        public OpenXmlFontRef(OpenXmlFontCollectionIndex index)
        {
            this.index = index;
            this.color = null;
        }

        public OpenXmlFontRef(OpenXmlFontCollectionIndex index, OpenXmlColor color)
        {
            this.index = index;
            this.color = color;
        }

        public OpenXmlFontCollectionIndex Index => this.index;

        public OpenXmlColor Color => this.color;
    }
}
