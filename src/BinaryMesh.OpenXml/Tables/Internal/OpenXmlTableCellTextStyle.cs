using System;
using DocumentFormat.OpenXml.Drawing;

using BinaryMesh.OpenXml.Shared;
using BinaryMesh.OpenXml.Styles;

namespace BinaryMesh.OpenXml.Tables.Internal
{
    internal class OpenXmlTableCellTextStyle : ITableCellTextStyle
    {
        private readonly ITablePartStyle result;

        private readonly ElementGenerator<TableCellTextStyle> tableCellTextStyleGenerator;

        public OpenXmlTableCellTextStyle(ITablePartStyle result, ElementGenerator<TableCellTextStyle> tableCellTextStyleGenerator)
        {
            this.result = result;
            this.tableCellTextStyleGenerator = tableCellTextStyleGenerator;
        }

        public ITablePartStyle SetFont(string typeface)
        {
            TableCellTextStyle tableCellTextStyle = this.tableCellTextStyleGenerator(true);
            tableCellTextStyle.RemoveAllChildren<FontReference>();
            tableCellTextStyle.RemoveAllChildren<Fonts>();

            
            tableCellTextStyle.PrependChild(
                new Fonts(
                    new LatinFont() { Typeface = typeface },
                    new ComplexScriptFont() { Typeface = typeface }
                )
            );

            return this.result;
        }

        public ITablePartStyle SetFont(OpenXmlFontRef fontRef)
        {
            TableCellTextStyle tableCellTextStyle = this.tableCellTextStyleGenerator(true);
            tableCellTextStyle.RemoveAllChildren<FontReference>();
            tableCellTextStyle.RemoveAllChildren<Fonts>();

            FontReference element = new FontReference
            {
                Index = (FontCollectionIndexValues)fontRef.Index
            };
            
            if (fontRef.Color != null)
            {
                element.AppendChild(fontRef.Color.CreateColorElement());
            }

            tableCellTextStyle.PrependChild(element);

            return this.result;
        }

        public ITablePartStyle SetFontColor(OpenXmlColor color)
        {
            TableCellTextStyle tableCellTextStyle = this.tableCellTextStyleGenerator(true);
            tableCellTextStyle.RemoveAllChildren<SchemeColor>();
            tableCellTextStyle.RemoveAllChildren<HslColor>();
            tableCellTextStyle.RemoveAllChildren<RgbColorModelHex>();
            tableCellTextStyle.RemoveAllChildren<RgbColorModelPercentage>();
            tableCellTextStyle.RemoveAllChildren<SystemColor>();
            tableCellTextStyle.RemoveAllChildren<PresetColor>();

            tableCellTextStyle.AppendChild(color.CreateColorElement());

            return this.result;
        }

        public ITablePartStyle SetIsBold(bool bold)
        {
            this.tableCellTextStyleGenerator(true).Bold = bold ? BooleanStyleValues.On : BooleanStyleValues.Off;
            return this.result;
        }

        public ITablePartStyle SetIsItalic(bool italic)
        {
            this.tableCellTextStyleGenerator(true).Italic = italic ? BooleanStyleValues.On : BooleanStyleValues.Off;
            return this.result;
        }
    }
}
