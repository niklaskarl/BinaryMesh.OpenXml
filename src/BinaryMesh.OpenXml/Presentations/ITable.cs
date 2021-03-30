using System;
using System.Collections.Generic;

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

    public interface ITableCell
    {
        IVisualStyle<ITableCell> Style { get; }

        ITextContent<ITableCell> Text { get; }
    }
}
