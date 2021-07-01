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

        ITableVisual SetStyle(ITableStyle style);

        ITableVisual SetHasFirstRow(bool value);

        ITableVisual SetHasFirstColumn(bool value);

        ITableVisual SetHasLastRow(bool value);

        ITableVisual SetHasLastColumn(bool value);

        ITableVisual SetHasBandRow(bool value);

        ITableVisual SetHasBandColumn(bool value);

        ITableVisual SetIsRightToLeft(bool value);

        ITableColumn AppendColumn(long width);

        ITableRow AppendRow(long height);
    }
}
