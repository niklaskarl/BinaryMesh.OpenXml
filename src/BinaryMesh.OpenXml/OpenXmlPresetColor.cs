using System;
using DocumentFormat.OpenXml;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace BinaryMesh.OpenXml
{
    public enum PresetColorValues
    {
        AliceBlue = Drawing.PresetColorValues.AliceBlue, // aliceBlue
        AntiqueWhite = Drawing.PresetColorValues.AntiqueWhite, // antiqueWhite
        Aqua = Drawing.PresetColorValues.Aqua, // aqua
        Aquamarine = Drawing.PresetColorValues.Aquamarine, // aquamarine
        Azure = Drawing.PresetColorValues.Azure, // azure
        Beige = Drawing.PresetColorValues.Beige, // beige
        Bisque = Drawing.PresetColorValues.Bisque, // bisque
        Black = Drawing.PresetColorValues.Black, // black
        BlanchedAlmond = Drawing.PresetColorValues.BlanchedAlmond, // blanchedAlmond
        Blue = Drawing.PresetColorValues.Blue, // blue
        BlueViolet = Drawing.PresetColorValues.BlueViolet, // blueViolet
        Brown = Drawing.PresetColorValues.Brown, // brown
        BurlyWood = Drawing.PresetColorValues.BurlyWood, // burlyWood
        CadetBlue = Drawing.PresetColorValues.CadetBlue, // cadetBlue
        Chartreuse = Drawing.PresetColorValues.Chartreuse, // chartreuse
        Chocolate = Drawing.PresetColorValues.Chocolate, // chocolate
        Coral = Drawing.PresetColorValues.Coral, // coral
        CornflowerBlue = Drawing.PresetColorValues.CornflowerBlue, // cornflowerBlue
        Cornsilk = Drawing.PresetColorValues.Cornsilk, // cornsilk
        Crimson = Drawing.PresetColorValues.Crimson, // crimson
        Cyan = Drawing.PresetColorValues.Cyan, // cyan
        DarkBlue = Drawing.PresetColorValues.DarkBlue, // darkBlue
        DarkCyan = Drawing.PresetColorValues.DarkCyan, // darkCyan
        DarkGoldenrod = Drawing.PresetColorValues.DarkGoldenrod, // darkGoldenrod
        DarkGray = Drawing.PresetColorValues.DarkGray, // darkGray
        DarkGrey = Drawing.PresetColorValues.DarkGrey, // darkGrey
        DarkGreen = Drawing.PresetColorValues.DarkGreen, // darkGreen
        DarkKhaki = Drawing.PresetColorValues.DarkKhaki, // darkKhaki
        DarkMagenta = Drawing.PresetColorValues.DarkMagenta, // darkMagenta
        DarkOliveGreen = Drawing.PresetColorValues.DarkOliveGreen, // darkOliveGreen
        DarkOrange = Drawing.PresetColorValues.DarkOrange, // darkOrange
        DarkOrchid = Drawing.PresetColorValues.DarkOrchid, // darkOrchid
        DarkRed = Drawing.PresetColorValues.DarkRed, // darkRed
        DarkSalmon = Drawing.PresetColorValues.DarkSalmon, // darkSalmon
        DarkSeaGreen = Drawing.PresetColorValues.DarkSeaGreen, // darkSeaGreen
        DarkSlateBlue = Drawing.PresetColorValues.DarkSlateBlue, // darkSlateBlue
        DarkSlateGray = Drawing.PresetColorValues.DarkSlateGray, // darkSlateGray
        DarkSlateGrey = Drawing.PresetColorValues.DarkSlateGrey, // darkSlateGrey
        DarkTurquoise = Drawing.PresetColorValues.DarkTurquoise, // darkTurquoise
        DarkViolet = Drawing.PresetColorValues.DarkViolet, // darkViolet
        DeepPink = Drawing.PresetColorValues.DeepPink, // deepPink
        DeepSkyBlue = Drawing.PresetColorValues.DeepSkyBlue, // deepSkyBlue
        DimGray = Drawing.PresetColorValues.DimGray, // dimGray
        DimGrey = Drawing.PresetColorValues.DimGrey, // dimGrey
        DodgerBlue = Drawing.PresetColorValues.DodgerBlue, // dodgerBlue
        Firebrick = Drawing.PresetColorValues.Firebrick, // firebrick
        FloralWhite = Drawing.PresetColorValues.FloralWhite, // floralWhite
        ForestGreen = Drawing.PresetColorValues.ForestGreen, // forestGreen
        Fuchsia = Drawing.PresetColorValues.Fuchsia, // fuchsia
        Gainsboro = Drawing.PresetColorValues.Gainsboro, // gainsboro
        GhostWhite = Drawing.PresetColorValues.GhostWhite, // ghostWhite
        Gold = Drawing.PresetColorValues.Gold, // gold
        Goldenrod = Drawing.PresetColorValues.Goldenrod, // goldenrod
        Gray = Drawing.PresetColorValues.Gray, // gray
        Grey = Drawing.PresetColorValues.Grey, // grey
        Green = Drawing.PresetColorValues.Green, // green
        GreenYellow = Drawing.PresetColorValues.GreenYellow, // greenYellow
        Honeydew = Drawing.PresetColorValues.Honeydew, // honeydew
        HotPink = Drawing.PresetColorValues.HotPink, // hotPink
        IndianRed = Drawing.PresetColorValues.IndianRed, // indianRed
        Indigo = Drawing.PresetColorValues.Indigo, // indigo
        Ivory = Drawing.PresetColorValues.Ivory, // ivory
        Khaki = Drawing.PresetColorValues.Khaki, // khaki
        Lavender = Drawing.PresetColorValues.Lavender, // lavender
        LavenderBlush = Drawing.PresetColorValues.LavenderBlush, // lavenderBlush
        LawnGreen = Drawing.PresetColorValues.LawnGreen, // lawnGreen
        LemonChiffon = Drawing.PresetColorValues.LemonChiffon, // lemonChiffon
        LightBlue = Drawing.PresetColorValues.LightBlue, // lightBlue
        LightCoral = Drawing.PresetColorValues.LightCoral, // lightCoral
        LightCyan = Drawing.PresetColorValues.LightCyan, // lightCyan
        LightGoldenrodYellow = Drawing.PresetColorValues.LightGoldenrodYellow, // lightGoldenrodYellow
        LightGray = Drawing.PresetColorValues.LightGray, // lightGray
        LightGrey = Drawing.PresetColorValues.LightGrey, // lightGrey
        LightGreen = Drawing.PresetColorValues.LightGreen, // lightGreen
        LightPink = Drawing.PresetColorValues.LightPink, // lightPink
        LightSalmon = Drawing.PresetColorValues.LightSalmon, // lightSalmon
        LightSeaGreen = Drawing.PresetColorValues.LightSeaGreen, // lightSeaGreen
        LightSkyBlue = Drawing.PresetColorValues.LightSkyBlue, // lightSkyBlue
        LightSlateGray = Drawing.PresetColorValues.LightSlateGray, // lightSlateGray
        LightSlateGrey = Drawing.PresetColorValues.LightSlateGrey, // lightSlateGrey
        LightSteelBlue = Drawing.PresetColorValues.LightSteelBlue, // lightSteelBlue
        LightYellow = Drawing.PresetColorValues.LightYellow, // lightYellow
        Lime = Drawing.PresetColorValues.Lime, // lime
        LimeGreen = Drawing.PresetColorValues.LimeGreen, // limeGreen
        Linen = Drawing.PresetColorValues.Linen, // linen
        Magenta = Drawing.PresetColorValues.Magenta, // magenta
        Maroon = Drawing.PresetColorValues.Maroon, // maroon
        MedAquamarine = Drawing.PresetColorValues.MedAquamarine, // medAquamarine
        MediumBlue = Drawing.PresetColorValues.MediumBlue, // mediumBlue
        MediumOrchid = Drawing.PresetColorValues.MediumOrchid, // mediumOrchid
        MediumPurple = Drawing.PresetColorValues.MediumPurple, // mediumPurple
        MediumSeaGreen = Drawing.PresetColorValues.MediumSeaGreen, // mediumSeaGreen
        MediumSlateBlue = Drawing.PresetColorValues.MediumSlateBlue, // mediumSlateBlue
        MediumSpringGreen = Drawing.PresetColorValues.MediumSpringGreen, // mediumSpringGreen
        MediumTurquoise = Drawing.PresetColorValues.MediumTurquoise, // mediumTurquoise
        MediumVioletRed = Drawing.PresetColorValues.MediumVioletRed, // mediumVioletRed
        MidnightBlue = Drawing.PresetColorValues.MidnightBlue, // midnightBlue
        MintCream = Drawing.PresetColorValues.MintCream, // mintCream
        MistyRose = Drawing.PresetColorValues.MistyRose, // mistyRose
        Moccasin = Drawing.PresetColorValues.Moccasin, // moccasin
        NavajoWhite = Drawing.PresetColorValues.NavajoWhite, // navajoWhite
        Navy = Drawing.PresetColorValues.Navy, // navy
        OldLace = Drawing.PresetColorValues.OldLace, // oldLace
        Olive = Drawing.PresetColorValues.Olive, // olive
        OliveDrab = Drawing.PresetColorValues.OliveDrab, // oliveDrab
        Orange = Drawing.PresetColorValues.Orange, // orange
        OrangeRed = Drawing.PresetColorValues.OrangeRed, // orangeRed
        Orchid = Drawing.PresetColorValues.Orchid, // orchid
        PaleGoldenrod = Drawing.PresetColorValues.PaleGoldenrod, // paleGoldenrod
        PaleGreen = Drawing.PresetColorValues.PaleGreen, // paleGreen
        PaleTurquoise = Drawing.PresetColorValues.PaleTurquoise, // paleTurquoise
        PaleVioletRed = Drawing.PresetColorValues.PaleVioletRed, // paleVioletRed
        PapayaWhip = Drawing.PresetColorValues.PapayaWhip, // papayaWhip
        PeachPuff = Drawing.PresetColorValues.PeachPuff, // peachPuff
        Peru = Drawing.PresetColorValues.Peru, // peru
        Pink = Drawing.PresetColorValues.Pink, // pink
        Plum = Drawing.PresetColorValues.Plum, // plum
        PowderBlue = Drawing.PresetColorValues.PowderBlue, // powderBlue
        Purple = Drawing.PresetColorValues.Purple, // purple
        Red = Drawing.PresetColorValues.Red, // red
        RosyBrown = Drawing.PresetColorValues.RosyBrown, // rosyBrown
        RoyalBlue = Drawing.PresetColorValues.RoyalBlue, // royalBlue
        SaddleBrown = Drawing.PresetColorValues.SaddleBrown, // saddleBrown
        Salmon = Drawing.PresetColorValues.Salmon, // salmon
        SandyBrown = Drawing.PresetColorValues.SandyBrown, // sandyBrown
        SeaGreen = Drawing.PresetColorValues.SeaGreen, // seaGreen
        SeaShell = Drawing.PresetColorValues.SeaShell, // seaShell
        Sienna = Drawing.PresetColorValues.Sienna, // sienna
        Silver = Drawing.PresetColorValues.Silver, // silver
        SkyBlue = Drawing.PresetColorValues.SkyBlue, // skyBlue
        SlateBlue = Drawing.PresetColorValues.SlateBlue, // slateBlue
        SlateGray = Drawing.PresetColorValues.SlateGray, // slateGray
        SlateGrey = Drawing.PresetColorValues.SlateGrey, // slateGrey
        Snow = Drawing.PresetColorValues.Snow, // snow
        SpringGreen = Drawing.PresetColorValues.SpringGreen, // springGreen
        SteelBlue = Drawing.PresetColorValues.SteelBlue, // steelBlue
        Tan = Drawing.PresetColorValues.Tan, // tan
        Teal = Drawing.PresetColorValues.Teal, // teal
        Thistle = Drawing.PresetColorValues.Thistle, // thistle
        Tomato = Drawing.PresetColorValues.Tomato, // tomato
        Turquoise = Drawing.PresetColorValues.Turquoise, // turquoise
        Violet = Drawing.PresetColorValues.Violet, // violet
        Wheat = Drawing.PresetColorValues.Wheat, // wheat
        White = Drawing.PresetColorValues.White, // white
        WhiteSmoke = Drawing.PresetColorValues.WhiteSmoke, // whiteSmoke
        Yellow = Drawing.PresetColorValues.Yellow, // yellow
        YellowGreen = Drawing.PresetColorValues.YellowGreen, // yellowGreen

    }

    public sealed class OpenXmlPresetColor : OpenXmlColor
    {
        private PresetColorValues color;

        public OpenXmlPresetColor(PresetColorValues color)
        {
            this.color = color;
        }

        internal OpenXmlPresetColor(OpenXmlPresetColor other)
            : base (other)
        {
            this.color = other.color;
        }
        
        internal override OpenXmlElement CreateColorElement()
        {
            Drawing.PresetColor element = new Drawing.PresetColor()
            {
                Val = (Drawing.PresetColorValues)this.color
            };

            this.AnnotateOpenXmlElement(element);

            return element;
        }

        internal override OpenXmlColor Clone()
        {
            return new OpenXmlPresetColor(this);
        }
    }
}
