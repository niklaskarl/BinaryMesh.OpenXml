using System;
using System.Collections.Generic;

using BinaryMesh.OpenXml.Styles;
using BinaryMesh.OpenXml.Tables;

namespace BinaryMesh.OpenXml.Presentations
{
    public interface ITableVisual : IVisual
    {
        new IVisualTransform<ITableVisual> Transform { get; }

        ITableCellCollection Cells { get; }

        IReadOnlyList<ITableColumn> Columns { get; }

        IReadOnlyList<ITableRow> Rows { get; }

        ITableColumn AppendColumn(long width);

        ITableRow AppendRow(long height);
    }
}
