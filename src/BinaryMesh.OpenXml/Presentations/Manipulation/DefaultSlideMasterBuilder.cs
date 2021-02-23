using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

using DocumentFormat.OpenXml.Presentation;

using Drawing = DocumentFormat.OpenXml.Drawing;

namespace BinaryMesh.OpenXml.Presentations
{
    internal static class DefaultSlideMasterBuilder
    {
        internal static SlideMaster BuildDefaultSlideMaster()
        {
            return new SlideMaster()
            {
                CommonSlideData = new CommonSlideData()
                {
                    Background = new Background()
                    {
                        BackgroundStyleReference = new BackgroundStyleReference()
                        {
                            Index = 1001,
                            SchemeColor = new Drawing.SchemeColor()
                            {
                                Val = Drawing.SchemeColorValues.Background1
                            }
                        }
                    },
                    ShapeTree = new ShapeTree()
                    {
                        NonVisualGroupShapeProperties = new NonVisualGroupShapeProperties()
                        {
                            NonVisualDrawingProperties = new NonVisualDrawingProperties()
                            {
                                Id = 1,
                                Name = ""
                            },
                            NonVisualGroupShapeDrawingProperties = new NonVisualGroupShapeDrawingProperties()
                            {
                            },
                            ApplicationNonVisualDrawingProperties = new ApplicationNonVisualDrawingProperties()
                            {
                            }
                        },
                        GroupShapeProperties = new GroupShapeProperties()
                        {
                            TransformGroup = new Drawing.TransformGroup()
                            {
                                Offset = new Drawing.Offset()
                                {
                                    X = 0,
                                    Y = 0
                                },
                                Extents = new Drawing.Extents()
                                {
                                    Cx = 0,
                                    Cy = 0
                                },
                                ChildOffset = new Drawing.ChildOffset()
                                {
                                    X = 0,
                                    Y = 0
                                },
                                ChildExtents = new Drawing.ChildExtents()
                                {
                                    Cx = 0,
                                    Cy = 0
                                }
                            }
                        },
                    }
                },
                ColorMap = new DocumentFormat.OpenXml.Presentation.ColorMap()
                {
                    Background1 = Drawing.ColorSchemeIndexValues.Light1,
                    Text1 = Drawing.ColorSchemeIndexValues.Light1,
                    Background2 = Drawing.ColorSchemeIndexValues.Light2,
                    Text2 = Drawing.ColorSchemeIndexValues.Dark2,
                    Accent1 = Drawing.ColorSchemeIndexValues.Accent1,
                    Accent2 = Drawing.ColorSchemeIndexValues.Accent2,
                    Accent3 = Drawing.ColorSchemeIndexValues.Accent3,
                    Accent4 = Drawing.ColorSchemeIndexValues.Accent4,
                    Accent5 = Drawing.ColorSchemeIndexValues.Accent5,
                    Accent6 = Drawing.ColorSchemeIndexValues.Accent6,
                    Hyperlink = Drawing.ColorSchemeIndexValues.Hyperlink,
                    FollowedHyperlink = Drawing.ColorSchemeIndexValues.FollowedHyperlink
                },
                TextStyles = new TextStyles()
                {
                    TitleStyle = new TitleStyle()
                    {
                        Level1ParagraphProperties = new Drawing.Level1ParagraphProperties()
                        {
                            Alignment = Drawing.TextAlignmentTypeValues.Left,
                            DefaultTabSize = 914400,
                            RightToLeft = false,
                            EastAsianLineBreak = true,
                            LatinLineBreak = false,
                            Height = true,
                            LineSpacing = new Drawing.LineSpacing() { SpacingPercent = new Drawing.SpacingPercent() { Val = 90000 } },
                            SpaceBefore = new Drawing.SpaceBefore() { SpacingPercent = new Drawing.SpacingPercent() { Val = 0 } }
                        }
                        .AppendChildFluent(new Drawing.NoBullet())
                        .AppendChildFluent(
                            new Drawing.DefaultRunProperties() { FontSize = 4400, Kerning = 1200 }
                            .AppendChildFluent(new Drawing.SolidFill() { SchemeColor = new Drawing.SchemeColor() { Val = Drawing.SchemeColorValues.Text1 } })
                            .AppendChildFluent(new Drawing.LatinFont() { Typeface = "+mj-lt" })
                            .AppendChildFluent(new Drawing.EastAsianFont() { Typeface = "+mj-ea" })
                            .AppendChildFluent(new Drawing.ComplexScriptFont() { Typeface = "+mj-cs" })
                        )
                    },
                    BodyStyle = new BodyStyle()
                    {
                        Level1ParagraphProperties = new Drawing.Level1ParagraphProperties()
                        {
                            LeftMargin = 228600,
                            Indent = -228600,
                            Alignment = Drawing.TextAlignmentTypeValues.Left,
                            DefaultTabSize = 914400,
                            RightToLeft = false,
                            EastAsianLineBreak = true,
                            LatinLineBreak = false,
                            Height = true,
                            LineSpacing = new Drawing.LineSpacing() { SpacingPercent = new Drawing.SpacingPercent() { Val = 90000 } },
                            SpaceBefore = new Drawing.SpaceBefore() { SpacingPercent = new Drawing.SpacingPercent() { Val = 1000 } }
                        }
                        .AppendChildFluent(new Drawing.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 })
                        .AppendChildFluent(new Drawing.CharacterBullet() { Char = "•" })
                        .AppendChildFluent(
                            new Drawing.DefaultRunProperties() { FontSize = 2800, Kerning = 1200 }
                            .AppendChildFluent(new Drawing.SolidFill() { SchemeColor = new Drawing.SchemeColor() { Val = Drawing.SchemeColorValues.Text1 } })
                            .AppendChildFluent(new Drawing.LatinFont() { Typeface = "+mn-lt" })
                            .AppendChildFluent(new Drawing.EastAsianFont() { Typeface = "+mn-ea" })
                            .AppendChildFluent(new Drawing.ComplexScriptFont() { Typeface = "+mn-cs" })
                        ),
                        Level2ParagraphProperties = new Drawing.Level2ParagraphProperties()
                        {
                            LeftMargin = 685800,
                            Indent = -228600,
                            Alignment = Drawing.TextAlignmentTypeValues.Left,
                            DefaultTabSize = 914400,
                            RightToLeft = false,
                            EastAsianLineBreak = true,
                            LatinLineBreak = false,
                            Height = true,
                            LineSpacing = new Drawing.LineSpacing() { SpacingPercent = new Drawing.SpacingPercent() { Val = 90000 } },
                            SpaceBefore = new Drawing.SpaceBefore() { SpacingPercent = new Drawing.SpacingPercent() { Val = 500 } }
                        }
                        .AppendChildFluent(new Drawing.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 })
                        .AppendChildFluent(new Drawing.CharacterBullet() { Char = "•" })
                        .AppendChildFluent(
                            new Drawing.DefaultRunProperties() { FontSize = 2400, Kerning = 1200 }
                            .AppendChildFluent(new Drawing.SolidFill() { SchemeColor = new Drawing.SchemeColor() { Val = Drawing.SchemeColorValues.Text1 } })
                            .AppendChildFluent(new Drawing.LatinFont() { Typeface = "+mn-lt" })
                            .AppendChildFluent(new Drawing.EastAsianFont() { Typeface = "+mn-ea" })
                            .AppendChildFluent(new Drawing.ComplexScriptFont() { Typeface = "+mn-cs" })
                        ),
                        Level3ParagraphProperties = new Drawing.Level3ParagraphProperties()
                        {
                            LeftMargin = 1143000,
                            Indent = -228600,
                            Alignment = Drawing.TextAlignmentTypeValues.Left,
                            DefaultTabSize = 914400,
                            RightToLeft = false,
                            EastAsianLineBreak = true,
                            LatinLineBreak = false,
                            Height = true,
                            LineSpacing = new Drawing.LineSpacing() { SpacingPercent = new Drawing.SpacingPercent() { Val = 90000 } },
                            SpaceBefore = new Drawing.SpaceBefore() { SpacingPercent = new Drawing.SpacingPercent() { Val = 500 } }
                        }
                        .AppendChildFluent(new Drawing.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 })
                        .AppendChildFluent(new Drawing.CharacterBullet() { Char = "•" })
                        .AppendChildFluent(
                            new Drawing.DefaultRunProperties() { FontSize = 2000, Kerning = 1200 }
                            .AppendChildFluent(new Drawing.SolidFill() { SchemeColor = new Drawing.SchemeColor() { Val = Drawing.SchemeColorValues.Text1 } })
                            .AppendChildFluent(new Drawing.LatinFont() { Typeface = "+mn-lt" })
                            .AppendChildFluent(new Drawing.EastAsianFont() { Typeface = "+mn-ea" })
                            .AppendChildFluent(new Drawing.ComplexScriptFont() { Typeface = "+mn-cs" })
                        ),
                        Level4ParagraphProperties = new Drawing.Level4ParagraphProperties()
                        {
                            LeftMargin = 1600200,
                            Indent = -228600,
                            Alignment = Drawing.TextAlignmentTypeValues.Left,
                            DefaultTabSize = 914400,
                            RightToLeft = false,
                            EastAsianLineBreak = true,
                            LatinLineBreak = false,
                            Height = true,
                            LineSpacing = new Drawing.LineSpacing() { SpacingPercent = new Drawing.SpacingPercent() { Val = 90000 } },
                            SpaceBefore = new Drawing.SpaceBefore() { SpacingPercent = new Drawing.SpacingPercent() { Val = 500 } }
                        }
                        .AppendChildFluent(new Drawing.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 })
                        .AppendChildFluent(new Drawing.CharacterBullet() { Char = "•" })
                        .AppendChildFluent(
                            new Drawing.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 }
                            .AppendChildFluent(new Drawing.SolidFill() { SchemeColor = new Drawing.SchemeColor() { Val = Drawing.SchemeColorValues.Text1 } })
                            .AppendChildFluent(new Drawing.LatinFont() { Typeface = "+mn-lt" })
                            .AppendChildFluent(new Drawing.EastAsianFont() { Typeface = "+mn-ea" })
                            .AppendChildFluent(new Drawing.ComplexScriptFont() { Typeface = "+mn-cs" })
                        ),
                        Level5ParagraphProperties = new Drawing.Level5ParagraphProperties()
                        {
                            LeftMargin = 2057400,
                            Indent = -228600,
                            Alignment = Drawing.TextAlignmentTypeValues.Left,
                            DefaultTabSize = 914400,
                            RightToLeft = false,
                            EastAsianLineBreak = true,
                            LatinLineBreak = false,
                            Height = true,
                            LineSpacing = new Drawing.LineSpacing() { SpacingPercent = new Drawing.SpacingPercent() { Val = 90000 } },
                            SpaceBefore = new Drawing.SpaceBefore() { SpacingPercent = new Drawing.SpacingPercent() { Val = 500 } }
                        }
                        .AppendChildFluent(new Drawing.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 })
                        .AppendChildFluent(new Drawing.CharacterBullet() { Char = "•" })
                        .AppendChildFluent(
                            new Drawing.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 }
                            .AppendChildFluent(new Drawing.SolidFill() { SchemeColor = new Drawing.SchemeColor() { Val = Drawing.SchemeColorValues.Text1 } })
                            .AppendChildFluent(new Drawing.LatinFont() { Typeface = "+mn-lt" })
                            .AppendChildFluent(new Drawing.EastAsianFont() { Typeface = "+mn-ea" })
                            .AppendChildFluent(new Drawing.ComplexScriptFont() { Typeface = "+mn-cs" })
                        ),
                        Level6ParagraphProperties = new Drawing.Level6ParagraphProperties()
                        {
                            LeftMargin = 2514600,
                            Indent = -228600,
                            Alignment = Drawing.TextAlignmentTypeValues.Left,
                            DefaultTabSize = 914400,
                            RightToLeft = false,
                            EastAsianLineBreak = true,
                            LatinLineBreak = false,
                            Height = true,
                            LineSpacing = new Drawing.LineSpacing() { SpacingPercent = new Drawing.SpacingPercent() { Val = 90000 } },
                            SpaceBefore = new Drawing.SpaceBefore() { SpacingPercent = new Drawing.SpacingPercent() { Val = 500 } }
                        }
                        .AppendChildFluent(new Drawing.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 })
                        .AppendChildFluent(new Drawing.CharacterBullet() { Char = "•" })
                        .AppendChildFluent(
                            new Drawing.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 }
                            .AppendChildFluent(new Drawing.SolidFill() { SchemeColor = new Drawing.SchemeColor() { Val = Drawing.SchemeColorValues.Text1 } })
                            .AppendChildFluent(new Drawing.LatinFont() { Typeface = "+mn-lt" })
                            .AppendChildFluent(new Drawing.EastAsianFont() { Typeface = "+mn-ea" })
                            .AppendChildFluent(new Drawing.ComplexScriptFont() { Typeface = "+mn-cs" })
                        ),
                        Level7ParagraphProperties = new Drawing.Level7ParagraphProperties()
                        {
                            LeftMargin = 2971800,
                            Indent = -228600,
                            Alignment = Drawing.TextAlignmentTypeValues.Left,
                            DefaultTabSize = 914400,
                            RightToLeft = false,
                            EastAsianLineBreak = true,
                            LatinLineBreak = false,
                            Height = true,
                            LineSpacing = new Drawing.LineSpacing() { SpacingPercent = new Drawing.SpacingPercent() { Val = 90000 } },
                            SpaceBefore = new Drawing.SpaceBefore() { SpacingPercent = new Drawing.SpacingPercent() { Val = 500 } }
                        }
                        .AppendChildFluent(new Drawing.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 })
                        .AppendChildFluent(new Drawing.CharacterBullet() { Char = "•" })
                        .AppendChildFluent(
                            new Drawing.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 }
                            .AppendChildFluent(new Drawing.SolidFill() { SchemeColor = new Drawing.SchemeColor() { Val = Drawing.SchemeColorValues.Text1 } })
                            .AppendChildFluent(new Drawing.LatinFont() { Typeface = "+mn-lt" })
                            .AppendChildFluent(new Drawing.EastAsianFont() { Typeface = "+mn-ea" })
                            .AppendChildFluent(new Drawing.ComplexScriptFont() { Typeface = "+mn-cs" })
                        ),
                        Level8ParagraphProperties = new Drawing.Level8ParagraphProperties()
                        {
                            LeftMargin = 3429000,
                            Indent = -228600,
                            Alignment = Drawing.TextAlignmentTypeValues.Left,
                            DefaultTabSize = 914400,
                            RightToLeft = false,
                            EastAsianLineBreak = true,
                            LatinLineBreak = false,
                            Height = true,
                            LineSpacing = new Drawing.LineSpacing() { SpacingPercent = new Drawing.SpacingPercent() { Val = 90000 } },
                            SpaceBefore = new Drawing.SpaceBefore() { SpacingPercent = new Drawing.SpacingPercent() { Val = 500 } }
                        }
                        .AppendChildFluent(new Drawing.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 })
                        .AppendChildFluent(new Drawing.CharacterBullet() { Char = "•" })
                        .AppendChildFluent(
                            new Drawing.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 }
                            .AppendChildFluent(new Drawing.SolidFill() { SchemeColor = new Drawing.SchemeColor() { Val = Drawing.SchemeColorValues.Text1 } })
                            .AppendChildFluent(new Drawing.LatinFont() { Typeface = "+mn-lt" })
                            .AppendChildFluent(new Drawing.EastAsianFont() { Typeface = "+mn-ea" })
                            .AppendChildFluent(new Drawing.ComplexScriptFont() { Typeface = "+mn-cs" })
                        ),
                        Level9ParagraphProperties = new Drawing.Level9ParagraphProperties()
                        {
                            LeftMargin = 3886200,
                            Indent = -228600,
                            Alignment = Drawing.TextAlignmentTypeValues.Left,
                            DefaultTabSize = 914400,
                            RightToLeft = false,
                            EastAsianLineBreak = true,
                            LatinLineBreak = false,
                            Height = true,
                            LineSpacing = new Drawing.LineSpacing() { SpacingPercent = new Drawing.SpacingPercent() { Val = 90000 } },
                            SpaceBefore = new Drawing.SpaceBefore() { SpacingPercent = new Drawing.SpacingPercent() { Val = 500 } }
                        }
                        .AppendChildFluent(new Drawing.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 })
                        .AppendChildFluent(new Drawing.CharacterBullet() { Char = "•" })
                        .AppendChildFluent(
                            new Drawing.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 }
                            .AppendChildFluent(new Drawing.SolidFill() { SchemeColor = new Drawing.SchemeColor() { Val = Drawing.SchemeColorValues.Text1 } })
                            .AppendChildFluent(new Drawing.LatinFont() { Typeface = "+mn-lt" })
                            .AppendChildFluent(new Drawing.EastAsianFont() { Typeface = "+mn-ea" })
                            .AppendChildFluent(new Drawing.ComplexScriptFont() { Typeface = "+mn-cs" })
                        )
                    },
                    OtherStyle = new OtherStyle()
                    {
                        DefaultParagraphProperties = new Drawing.DefaultParagraphProperties()
                        .AppendChildFluent(new Drawing.DefaultRunProperties() { Language = "de-De" }),
                        Level1ParagraphProperties = new Drawing.Level1ParagraphProperties()
                        {
                            LeftMargin = 0,
                            Alignment = Drawing.TextAlignmentTypeValues.Left,
                            DefaultTabSize = 914400,
                            RightToLeft = false,
                            EastAsianLineBreak = true,
                            LatinLineBreak = false,
                            Height = true
                        }
                        .AppendChildFluent(
                            new Drawing.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 }
                            .AppendChildFluent(new Drawing.SolidFill() { SchemeColor = new Drawing.SchemeColor() { Val = Drawing.SchemeColorValues.Text1 } })
                            .AppendChildFluent(new Drawing.LatinFont() { Typeface = "+mn-lt" })
                            .AppendChildFluent(new Drawing.EastAsianFont() { Typeface = "+mn-ea" })
                            .AppendChildFluent(new Drawing.ComplexScriptFont() { Typeface = "+mn-cs" })
                        ),
                        Level2ParagraphProperties = new Drawing.Level2ParagraphProperties()
                        {
                            LeftMargin = 457200,
                            Alignment = Drawing.TextAlignmentTypeValues.Left,
                            DefaultTabSize = 914400,
                            RightToLeft = false,
                            EastAsianLineBreak = true,
                            LatinLineBreak = false,
                            Height = true
                        }
                        .AppendChildFluent(
                            new Drawing.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 }
                            .AppendChildFluent(new Drawing.SolidFill() { SchemeColor = new Drawing.SchemeColor() { Val = Drawing.SchemeColorValues.Text1 } })
                            .AppendChildFluent(new Drawing.LatinFont() { Typeface = "+mn-lt" })
                            .AppendChildFluent(new Drawing.EastAsianFont() { Typeface = "+mn-ea" })
                            .AppendChildFluent(new Drawing.ComplexScriptFont() { Typeface = "+mn-cs" })
                        ),
                        Level3ParagraphProperties = new Drawing.Level3ParagraphProperties()
                        {
                            LeftMargin = 914400,
                            Alignment = Drawing.TextAlignmentTypeValues.Left,
                            DefaultTabSize = 914400,
                            RightToLeft = false,
                            EastAsianLineBreak = true,
                            LatinLineBreak = false,
                            Height = true
                        }
                        .AppendChildFluent(
                            new Drawing.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 }
                            .AppendChildFluent(new Drawing.SolidFill() { SchemeColor = new Drawing.SchemeColor() { Val = Drawing.SchemeColorValues.Text1 } })
                            .AppendChildFluent(new Drawing.LatinFont() { Typeface = "+mn-lt" })
                            .AppendChildFluent(new Drawing.EastAsianFont() { Typeface = "+mn-ea" })
                            .AppendChildFluent(new Drawing.ComplexScriptFont() { Typeface = "+mn-cs" })
                        ),
                        Level4ParagraphProperties = new Drawing.Level4ParagraphProperties()
                        {
                            LeftMargin = 1371600,
                            Alignment = Drawing.TextAlignmentTypeValues.Left,
                            DefaultTabSize = 914400,
                            RightToLeft = false,
                            EastAsianLineBreak = true,
                            LatinLineBreak = false,
                            Height = true
                        }
                        .AppendChildFluent(
                            new Drawing.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 }
                            .AppendChildFluent(new Drawing.SolidFill() { SchemeColor = new Drawing.SchemeColor() { Val = Drawing.SchemeColorValues.Text1 } })
                            .AppendChildFluent(new Drawing.LatinFont() { Typeface = "+mn-lt" })
                            .AppendChildFluent(new Drawing.EastAsianFont() { Typeface = "+mn-ea" })
                            .AppendChildFluent(new Drawing.ComplexScriptFont() { Typeface = "+mn-cs" })
                        ),
                        Level5ParagraphProperties = new Drawing.Level5ParagraphProperties()
                        {
                            LeftMargin = 1828800,
                            Alignment = Drawing.TextAlignmentTypeValues.Left,
                            DefaultTabSize = 914400,
                            RightToLeft = false,
                            EastAsianLineBreak = true,
                            LatinLineBreak = false,
                            Height = true
                        }
                        .AppendChildFluent(
                            new Drawing.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 }
                            .AppendChildFluent(new Drawing.SolidFill() { SchemeColor = new Drawing.SchemeColor() { Val = Drawing.SchemeColorValues.Text1 } })
                            .AppendChildFluent(new Drawing.LatinFont() { Typeface = "+mn-lt" })
                            .AppendChildFluent(new Drawing.EastAsianFont() { Typeface = "+mn-ea" })
                            .AppendChildFluent(new Drawing.ComplexScriptFont() { Typeface = "+mn-cs" })
                        ),
                        Level6ParagraphProperties = new Drawing.Level6ParagraphProperties()
                        {
                            LeftMargin = 2286000,
                            Alignment = Drawing.TextAlignmentTypeValues.Left,
                            DefaultTabSize = 914400,
                            RightToLeft = false,
                            EastAsianLineBreak = true,
                            LatinLineBreak = false,
                            Height = true
                        }
                        .AppendChildFluent(
                            new Drawing.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 }
                            .AppendChildFluent(new Drawing.SolidFill() { SchemeColor = new Drawing.SchemeColor() { Val = Drawing.SchemeColorValues.Text1 } })
                            .AppendChildFluent(new Drawing.LatinFont() { Typeface = "+mn-lt" })
                            .AppendChildFluent(new Drawing.EastAsianFont() { Typeface = "+mn-ea" })
                            .AppendChildFluent(new Drawing.ComplexScriptFont() { Typeface = "+mn-cs" })
                        ),
                        Level7ParagraphProperties = new Drawing.Level7ParagraphProperties()
                        {
                            LeftMargin = 2743200,
                            Alignment = Drawing.TextAlignmentTypeValues.Left,
                            DefaultTabSize = 914400,
                            RightToLeft = false,
                            EastAsianLineBreak = true,
                            LatinLineBreak = false,
                            Height = true
                        }
                        .AppendChildFluent(
                            new Drawing.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 }
                            .AppendChildFluent(new Drawing.SolidFill() { SchemeColor = new Drawing.SchemeColor() { Val = Drawing.SchemeColorValues.Text1 } })
                            .AppendChildFluent(new Drawing.LatinFont() { Typeface = "+mn-lt" })
                            .AppendChildFluent(new Drawing.EastAsianFont() { Typeface = "+mn-ea" })
                            .AppendChildFluent(new Drawing.ComplexScriptFont() { Typeface = "+mn-cs" })
                        ),
                        Level8ParagraphProperties = new Drawing.Level8ParagraphProperties()
                        {
                            LeftMargin = 3200400,
                            Alignment = Drawing.TextAlignmentTypeValues.Left,
                            DefaultTabSize = 914400,
                            RightToLeft = false,
                            EastAsianLineBreak = true,
                            LatinLineBreak = false,
                            Height = true
                        }
                        .AppendChildFluent(
                            new Drawing.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 }
                            .AppendChildFluent(new Drawing.SolidFill() { SchemeColor = new Drawing.SchemeColor() { Val = Drawing.SchemeColorValues.Text1 } })
                            .AppendChildFluent(new Drawing.LatinFont() { Typeface = "+mn-lt" })
                            .AppendChildFluent(new Drawing.EastAsianFont() { Typeface = "+mn-ea" })
                            .AppendChildFluent(new Drawing.ComplexScriptFont() { Typeface = "+mn-cs" })
                        ),
                        Level9ParagraphProperties = new Drawing.Level9ParagraphProperties()
                        {
                            LeftMargin = 3657600,
                            Alignment = Drawing.TextAlignmentTypeValues.Left,
                            DefaultTabSize = 914400,
                            RightToLeft = false,
                            EastAsianLineBreak = true,
                            LatinLineBreak = false,
                            Height = true
                        }
                        .AppendChildFluent(
                            new Drawing.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 }
                            .AppendChildFluent(new Drawing.SolidFill() { SchemeColor = new Drawing.SchemeColor() { Val = Drawing.SchemeColorValues.Text1 } })
                            .AppendChildFluent(new Drawing.LatinFont() { Typeface = "+mn-lt" })
                            .AppendChildFluent(new Drawing.EastAsianFont() { Typeface = "+mn-ea" })
                            .AppendChildFluent(new Drawing.ComplexScriptFont() { Typeface = "+mn-cs" })
                        )
                    }
                }
            };
        }

        internal static SlideLayout[] BuildDefaultSlideLayouts()
        {
            return new SlideLayout[]
            {
                BuildDefaultSlideLayout1()
            };
        }

        internal static SlideLayout BuildDefaultSlideLayout1()
        {
            return new SlideLayout()
            {
                CommonSlideData = new CommonSlideData()
                {
                    ShapeTree = new ShapeTree()
                    {
                        NonVisualGroupShapeProperties = new NonVisualGroupShapeProperties()
                        {
                            NonVisualDrawingProperties = new NonVisualDrawingProperties() { Id = 1, Name = "" },
                            NonVisualGroupShapeDrawingProperties = new NonVisualGroupShapeDrawingProperties(),
                            ApplicationNonVisualDrawingProperties = new ApplicationNonVisualDrawingProperties()
                        },
                        GroupShapeProperties = new GroupShapeProperties()
                        {
                            TransformGroup = new Drawing.TransformGroup()
                            {
                                Offset = new Drawing.Offset() { X = 0, Y = 0 },
                                Extents = new Drawing.Extents() { Cx = 0, Cy = 0 },
                                ChildOffset = new Drawing.ChildOffset() { X = 0, Y = 0 },
                                ChildExtents = new Drawing.ChildExtents() { Cx = 0, Cy = 0 }
                            }
                        },
                    }
                    .AppendChildFluent(
                        new Shape()
                        {
                            NonVisualShapeProperties = new NonVisualShapeProperties()
                            {
                                NonVisualDrawingProperties = new NonVisualDrawingProperties()
                                {
                                    Id = 2,
                                    Name = "Titel 1",
                                    /*NonVisualDrawingPropertiesExtensionList = new Drawing.NonVisualDrawingPropertiesExtensionList()
                                    .AppendChildFluent(new Drawing.Extension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" })*/
                                },
                                NonVisualShapeDrawingProperties = new NonVisualShapeDrawingProperties()
                                {
                                    ShapeLocks = new Drawing.ShapeLocks() { NoGrouping = true }
                                },
                                ApplicationNonVisualDrawingProperties = new ApplicationNonVisualDrawingProperties()
                                {
                                    PlaceholderShape = new PlaceholderShape() { Type = PlaceholderValues.Title }
                                }
                            },
                            ShapeProperties = new ShapeProperties()
                            {
                                Transform2D = new Drawing.Transform2D()
                                {
                                    Offset = new Drawing.Offset() { X = 838200, Y = 365125 },
                                    Extents = new Drawing.Extents() { Cx = 10515600, Cy = 1325563 }
                                },
                            }
                            .AppendChildFluent(new Drawing.PresetGeometry() { Preset = Drawing.ShapeTypeValues.Rectangle, AdjustValueList = new Drawing.AdjustValueList() }),
                            TextBody = new TextBody()
                            {
                                BodyProperties = new Drawing.BodyProperties(),
                                ListStyle = new Drawing.ListStyle(),
                            }
                            .AppendChildFluent(
                                new Drawing.Paragraph().AppendChildFluent(new Drawing.Run() { RunProperties = new Drawing.RunProperties() { Language = "de-DE" }, Text = new Drawing.Text() { Text = "Mastertitelformat bearbeiten" } })
                            )
                        }
                    )
                    .AppendChildFluent(
                        new Shape()
                        {
                            NonVisualShapeProperties = new NonVisualShapeProperties()
                            {
                                NonVisualDrawingProperties = new NonVisualDrawingProperties()
                                {
                                    Id = 3,
                                    Name = "Textplatzhalter 2",
                                    /*NonVisualDrawingPropertiesExtensionList = new Drawing.NonVisualDrawingPropertiesExtensionList()
                                    .AppendChildFluent(new Drawing.Extension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" })*/
                                },
                                NonVisualShapeDrawingProperties = new NonVisualShapeDrawingProperties()
                                {
                                    ShapeLocks = new Drawing.ShapeLocks() { NoGrouping = true }
                                },
                                ApplicationNonVisualDrawingProperties = new ApplicationNonVisualDrawingProperties()
                                {
                                    PlaceholderShape = new PlaceholderShape() { Type = PlaceholderValues.Body, Index = 1 }
                                }
                            },
                            ShapeProperties = new ShapeProperties()
                            {
                                Transform2D = new Drawing.Transform2D()
                                {
                                    Offset = new Drawing.Offset() { X = 838200, Y = 1825625 },
                                    Extents = new Drawing.Extents() { Cx = 10515600, Cy = 4351338 }
                                },
                            }
                            .AppendChildFluent(new Drawing.PresetGeometry() { Preset = Drawing.ShapeTypeValues.Rectangle, AdjustValueList = new Drawing.AdjustValueList() }),
                            TextBody = new TextBody()
                            {
                                BodyProperties = new Drawing.BodyProperties(),
                                ListStyle = new Drawing.ListStyle(),
                            }
                            .AppendChildFluent(
                                new Drawing.Paragraph() { ParagraphProperties = new Drawing.ParagraphProperties() { Level= 0 } }
                                    .AppendChildFluent(new Drawing.Run() { RunProperties = new Drawing.RunProperties() { Language = "de-DE" }, Text = new Drawing.Text() { Text = "Mastertextformat bearbeiten" } })
                            )
                            .AppendChildFluent(
                                new Drawing.Paragraph() { ParagraphProperties = new Drawing.ParagraphProperties() { Level= 1 } }
                                    .AppendChildFluent(new Drawing.Run() { RunProperties = new Drawing.RunProperties() { Language = "de-DE" }, Text = new Drawing.Text() { Text = "Zweite Ebene" } })
                            )
                            .AppendChildFluent(
                                new Drawing.Paragraph() { ParagraphProperties = new Drawing.ParagraphProperties() { Level= 2 } }
                                    .AppendChildFluent(new Drawing.Run() { RunProperties = new Drawing.RunProperties() { Language = "de-DE" }, Text = new Drawing.Text() { Text = "Dritte Ebene" } })
                            )
                            .AppendChildFluent(
                                new Drawing.Paragraph() { ParagraphProperties = new Drawing.ParagraphProperties() { Level= 3 } }
                                    .AppendChildFluent(new Drawing.Run() { RunProperties = new Drawing.RunProperties() { Language = "de-DE" }, Text = new Drawing.Text() { Text = "Vierte Ebene" } })
                            )
                            .AppendChildFluent(
                                new Drawing.Paragraph() { ParagraphProperties = new Drawing.ParagraphProperties() { Level= 4 } }
                                    .AppendChildFluent(new Drawing.Run() { RunProperties = new Drawing.RunProperties() { Language = "de-DE" }, Text = new Drawing.Text() { Text = "Fünfte Ebene" } })
                            )
                        }
                    )
                }
            };
        }
    }
}
