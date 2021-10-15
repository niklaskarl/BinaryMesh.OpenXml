using System;
using DocumentFormat.OpenXml.Packaging;

namespace BinaryMesh.OpenXml.Internal
{
    internal sealed class OpenXmlTheme : IOpenXmlTheme
    {
        private readonly IOpenXmlDocument document;

        private readonly ThemePart themePart;

        public OpenXmlTheme(IOpenXmlDocument document, ThemePart themePart)
        {
            this.document = document;
            this.themePart = themePart;
        }

        public ThemePart ThemePart => this.themePart;
        
        public string ResolveFontTypeface(string typeface)
        {
            switch (typeface)
            {
                case "+mj-lt":
                    typeface = this.themePart.Theme.ThemeElements?.FontScheme?.MajorFont?.LatinFont?.Typeface?.Value;
                    break;
                case "+mn-lt":
                    typeface = this.themePart.Theme.ThemeElements?.FontScheme?.MinorFont?.LatinFont?.Typeface?.Value;
                    break;
            }

            return typeface;
        }
    }
}
