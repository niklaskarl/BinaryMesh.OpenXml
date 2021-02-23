using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

using DocumentFormat.OpenXml.Drawing;

namespace BinaryMesh.OpenXml.Presentations
{
    internal static class DefaultThemeBuilder
    {
        internal static Theme BuildDefaultTheme()
        {
            Theme theme = new Theme()
            {
                ThemeElements = new ThemeElements()
                {
                    ColorScheme = new ColorScheme()
                    {
                        Name = "Office",
                        Dark1Color = new Dark1Color() { SystemColor = new SystemColor() { Val = SystemColorValues.WindowText, LastColor = "000000" } },
                        Light1Color = new Light1Color() { SystemColor = new SystemColor() { Val = SystemColorValues.Window, LastColor = "FFFFFF" } },
                        Dark2Color = new Dark2Color() { RgbColorModelHex = new RgbColorModelHex() { Val = "44546A" } },
                        Light2Color = new Light2Color() { RgbColorModelHex = new RgbColorModelHex() { Val = "E7E6E6" } },
                        Accent1Color = new Accent1Color() { RgbColorModelHex = new RgbColorModelHex() { Val = "4472C4" } },
                        Accent2Color = new Accent2Color() { RgbColorModelHex = new RgbColorModelHex() { Val = "ED7D31" } },
                        Accent3Color = new Accent3Color() { RgbColorModelHex = new RgbColorModelHex() { Val = "A5A5A5" } },
                        Accent4Color = new Accent4Color() { RgbColorModelHex = new RgbColorModelHex() { Val = "FFC000" } },
                        Accent5Color = new Accent5Color() { RgbColorModelHex = new RgbColorModelHex() { Val = "5B9BD5" } },
                        Accent6Color = new Accent6Color() { RgbColorModelHex = new RgbColorModelHex() { Val = "70AD47" } },
                        Hyperlink = new Hyperlink() { RgbColorModelHex = new RgbColorModelHex() { Val = "0563C1" } },
                        FollowedHyperlinkColor = new FollowedHyperlinkColor() { RgbColorModelHex = new RgbColorModelHex() { Val = "954F72" } }
                    },
                    FontScheme = new FontScheme()
                    {
                        Name = "Office",
                        MajorFont = new MajorFont()
                        {
                            LatinFont = new LatinFont() { Typeface = "Calibri Light", Panose = "020F0302020204030204" },
                            EastAsianFont = new EastAsianFont() { Typeface = "" },
                            ComplexScriptFont = new ComplexScriptFont() { Typeface = "" }
                            // TODO: fonts
                        },
                        MinorFont = new MinorFont()
                        {
                            LatinFont = new LatinFont() { Typeface = "Calibri", Panose = "020F0302020204030204" },
                            EastAsianFont = new EastAsianFont() { Typeface = "" },
                            ComplexScriptFont = new ComplexScriptFont() { Typeface = "" }
                            // TODO: fonts
                        }
                    },
                    FormatScheme = new FormatScheme()
                    {
                        Name = "Office",
                        FillStyleList = new FillStyleList()
                            .AppendChildFluent(new SolidFill() { SchemeColor = new SchemeColor() { Val = SchemeColorValues.PhColor } })
                            .AppendChildFluent(
                                new GradientFill()
                                {
                                    RotateWithShape = true,
                                    GradientStopList = new GradientStopList()
                                    .AppendChildFluent(
                                        new GradientStop()
                                        {
                                            Position = 0,
                                            SchemeColor = new SchemeColor() { Val = SchemeColorValues.PhColor }
                                            .AppendChildFluent(new LuminanceModulation() { Val = 110000 })
                                            .AppendChildFluent(new SaturationModulation() { Val = 105000 })
                                            .AppendChildFluent(new Tint() { Val = 67000 })
                                        }
                                    )
                                    .AppendChildFluent(
                                        new GradientStop()
                                        {
                                            Position = 50000,
                                            SchemeColor = new SchemeColor() { Val = SchemeColorValues.PhColor }
                                            .AppendChildFluent(new LuminanceModulation() { Val = 105000 })
                                            .AppendChildFluent(new SaturationModulation() { Val = 103000 })
                                            .AppendChildFluent(new Tint() { Val = 73000 })
                                        }
                                    )
                                    .AppendChildFluent(
                                        new GradientStop()
                                        {
                                            Position = 100000,
                                            SchemeColor = new SchemeColor() { Val = SchemeColorValues.PhColor }
                                            .AppendChildFluent(new LuminanceModulation() { Val = 105000 })
                                            .AppendChildFluent(new SaturationModulation() { Val = 109000 })
                                            .AppendChildFluent(new Tint() { Val = 81000 })
                                        }
                                    )
                                }
                                .AppendChildFluent(new LinearGradientFill() { Angle = 5400000, Scaled = false })
                            )
                            .AppendChildFluent(
                                new GradientFill()
                                {
                                    RotateWithShape = true,
                                    GradientStopList = new GradientStopList()
                                    .AppendChildFluent(
                                        new GradientStop()
                                        {
                                            Position = 0,
                                            SchemeColor = new SchemeColor() { Val = SchemeColorValues.PhColor }
                                            .AppendChildFluent(new LuminanceModulation() { Val = 102000 })
                                            .AppendChildFluent(new SaturationModulation() { Val = 103000 })
                                            .AppendChildFluent(new Tint() { Val = 94000 })
                                        }
                                    )
                                    .AppendChildFluent(
                                        new GradientStop()
                                        {
                                            Position = 50000,
                                            SchemeColor = new SchemeColor() { Val = SchemeColorValues.PhColor }
                                            .AppendChildFluent(new LuminanceModulation() { Val = 100000 })
                                            .AppendChildFluent(new SaturationModulation() { Val = 110000 })
                                            .AppendChildFluent(new Shade() { Val = 100000 })
                                        }
                                    )
                                    .AppendChildFluent(
                                        new GradientStop()
                                        {
                                            Position = 100000,
                                            SchemeColor = new SchemeColor() { Val = SchemeColorValues.PhColor }
                                            .AppendChildFluent(new LuminanceModulation() { Val = 99000 })
                                            .AppendChildFluent(new SaturationModulation() { Val = 120000 })
                                            .AppendChildFluent(new Shade() { Val = 78000 })
                                        }
                                    )
                                }
                                .AppendChildFluent(new LinearGradientFill() { Angle = 5400000, Scaled = false })
                            ),
                        LineStyleList = new LineStyleList()
                            .AppendChildFluent(
                                new Outline() { Width = 6350, CapType = LineCapValues.Flat, CompoundLineType = CompoundLineValues.Single, Alignment = PenAlignmentValues.Center }
                                .AppendChildFluent(new SolidFill() { SchemeColor = new SchemeColor() { Val = SchemeColorValues.PhColor } })
                                .AppendChildFluent(new PresetDash() { Val = PresetLineDashValues.Solid })
                                .AppendChildFluent(new Miter() { Limit = 800000 })
                            )
                            .AppendChildFluent(
                                new Outline() { Width = 12700, CapType = LineCapValues.Flat, CompoundLineType = CompoundLineValues.Single, Alignment = PenAlignmentValues.Center }
                                .AppendChildFluent(new SolidFill() { SchemeColor = new SchemeColor() { Val = SchemeColorValues.PhColor } })
                                .AppendChildFluent(new PresetDash() { Val = PresetLineDashValues.Solid })
                                .AppendChildFluent(new Miter() { Limit = 800000 })
                            )
                            .AppendChildFluent(
                                new Outline() { Width = 19050, CapType = LineCapValues.Flat, CompoundLineType = CompoundLineValues.Single, Alignment = PenAlignmentValues.Center }
                                .AppendChildFluent(new SolidFill() { SchemeColor = new SchemeColor() { Val = SchemeColorValues.PhColor } })
                                .AppendChildFluent(new PresetDash() { Val = PresetLineDashValues.Solid })
                                .AppendChildFluent(new Miter() { Limit = 800000 })
                            ),
                        EffectStyleList = new EffectStyleList()
                            .AppendChildFluent(new EffectStyle().AppendChildFluent(new EffectList()))
                            .AppendChildFluent(new EffectStyle().AppendChildFluent(new EffectList()))
                            .AppendChildFluent(
                                new EffectStyle()
                                .AppendChildFluent(
                                    new EffectList()
                                    {
                                        OuterShadow = new OuterShadow()
                                        {
                                            BlurRadius = 57150,
                                            Distance = 19050,
                                            Direction = 5400000,
                                            Alignment = RectangleAlignmentValues.Center,
                                            RotateWithShape = false,
                                            RgbColorModelHex = new RgbColorModelHex()
                                            {
                                                Val = "000000",
                                            }
                                            .AppendChildFluent(new Alpha() { Val = 63000 })
                                        }
                                    }
                                )
                            ),
                        BackgroundFillStyleList = new BackgroundFillStyleList()
                            .AppendChildFluent(new SolidFill() { SchemeColor = new SchemeColor() { Val = SchemeColorValues.PhColor } })
                            .AppendChildFluent(
                                new SolidFill()
                                {
                                    SchemeColor = new SchemeColor()
                                    {
                                        Val = SchemeColorValues.PhColor
                                    }
                                    .AppendChildFluent(new SaturationModulation() { Val = 170000 })
                                    .AppendChildFluent(new Tint() { Val = 95000 })
                                }
                            )
                            .AppendChildFluent(
                                new GradientFill()
                                {
                                    RotateWithShape = true,
                                    GradientStopList = new GradientStopList()
                                    .AppendChildFluent(
                                        new GradientStop()
                                        {
                                            Position = 0,
                                            SchemeColor = new SchemeColor() { Val = SchemeColorValues.PhColor }
                                            .AppendChildFluent(new LuminanceModulation() { Val = 102000 })
                                            .AppendChildFluent(new SaturationModulation() { Val = 150000 })
                                            .AppendChildFluent(new Tint() { Val = 93000 })
                                            .AppendChildFluent(new Shade() { Val = 98000 })
                                        }
                                    )
                                    .AppendChildFluent(
                                        new GradientStop()
                                        {
                                            Position = 50000,
                                            SchemeColor = new SchemeColor() { Val = SchemeColorValues.PhColor }
                                            .AppendChildFluent(new LuminanceModulation() { Val = 103000 })
                                            .AppendChildFluent(new SaturationModulation() { Val = 130000 })
                                            .AppendChildFluent(new Tint() { Val = 98000 })
                                            .AppendChildFluent(new Shade() { Val = 90000 })
                                        }
                                    )
                                    .AppendChildFluent(
                                        new GradientStop()
                                        {
                                            Position = 100000,
                                            SchemeColor = new SchemeColor() { Val = SchemeColorValues.PhColor }
                                            .AppendChildFluent(new SaturationModulation() { Val = 120000 })
                                            .AppendChildFluent(new Shade() { Val = 63000 })
                                        }
                                    ),
                                }
                                .AppendChildFluent(new LinearGradientFill() { Angle = 5400000, Scaled = false })
                            )
                    }
                },
                ObjectDefaults = new ObjectDefaults(),
                ExtraColorSchemeList = new ExtraColorSchemeList(),

                OfficeStyleSheetExtensionList = new OfficeStyleSheetExtensionList()
                .AppendChildFluent(
                    new Extension()
                    {
                        Uri = "{05A4C25C-085E-4340-85A3-A5531E510DB2}"
                    }
                )
            };

            return theme;
        }

        internal static TableStyleList BuildDefaultTableStyleList()
        {
            TableStyleList tableStyleList = new TableStyleList()
            {
                Default = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"
            };

            return tableStyleList;
        }
    }
}
