using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;

namespace BinaryMesh.OpenXml.Internal
{
    internal sealed class OpenXmlParagraphTextStyle : IOpenXmlParagraphTextStyle
    {
        private readonly TextParagraphPropertiesType properties;

        public OpenXmlParagraphTextStyle(TextParagraphPropertiesType properties)
        {
            this.properties = properties;
        }

        public double? Size
        {
            get
            {
                Int32Value value = this.properties.GetFirstChild<DefaultRunProperties>()?.FontSize;
                return (value?.HasValue ?? false) ? value.Value / 100.0 : (double?)null;
            }
        }

        public double? Kerning
        {
            get
            {
                Int32Value value = this.properties.GetFirstChild<DefaultRunProperties>()?.Kerning;
                return (value?.HasValue ?? false) ? value.Value / 100.0 : (double?)null;
            }
        }

        public string LatinTypeface
        {
            get
            {
                StringValue value = this.properties.GetFirstChild<DefaultRunProperties>()?.GetFirstChild<LatinFont>()?.Typeface;
                return (value?.HasValue ?? false) ? value.Value : null;
            }
        }
    }
}
