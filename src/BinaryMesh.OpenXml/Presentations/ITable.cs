using System;
using System.Collections.Generic;

namespace BinaryMesh.OpenXml.Presentations
{
    public interface ITableVisual : IVisual
    {
        new ITableVisual SetOffset(long x, long y);

        new ITableVisual SetExtents(long width, long height);

        ITableCellCollection Cells { get; }

        IReadOnlyList<ITableColumn> Columns { get; }

        IReadOnlyList<ITableRow> Rows { get; }

        ITableColumn AppendColumn(long width);

        ITableRow AppendRow(long height);
    }

    public interface ITableCellCollection
    {
        ITableCell this[int column, int row] { get; }
    }

    public interface ITableColumn
    {
    }

    public interface ITableRow
    {
    }

    public interface ITableCell : ITextShape<ITableCell>
    {
    }
}
