using System;
using System.Collections.Generic;

using BinaryMesh.OpenXml.Styles;

namespace BinaryMesh.OpenXml.Tables
{
    public interface ITableCellCollection
    {
        ITableCell this[int column, int row] { get; }
    }
}
