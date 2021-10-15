using System;

namespace BinaryMesh.OpenXml.Internal
{
    internal interface IOpenXmlTextStyle
    {
        IOpenXmlParagraphTextStyle GetParagraphTextStyle(int level);
    }

    internal interface IOpenXmlParagraphTextStyle
    {
        double? Size { get; }
        
        double? Kerning { get; }
        
        string LatinTypeface { get; }
    }
}
