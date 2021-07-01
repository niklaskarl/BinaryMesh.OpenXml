using System;

using BinaryMesh.OpenXml.Styles;

namespace BinaryMesh.OpenXml.Tables
{
    public interface ITableCellBoderStyle
    {
        IStrokeStyle<ITablePartStyle> Left { get; }

        IStrokeStyle<ITablePartStyle> Top { get; }

        IStrokeStyle<ITablePartStyle> Right { get; }

        IStrokeStyle<ITablePartStyle> Bottom { get; }

        IStrokeStyle<ITablePartStyle> InsideHorizontal { get; }

        IStrokeStyle<ITablePartStyle> InsideVertical { get; }
    }
}
