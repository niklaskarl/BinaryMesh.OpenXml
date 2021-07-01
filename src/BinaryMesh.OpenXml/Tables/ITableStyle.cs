using System;

using BinaryMesh.OpenXml.Styles;

namespace BinaryMesh.OpenXml.Tables
{
    public interface ITableStyle
    {
        string Id { get; }

        ITablePartStyle WholeTablePart { get; }

        ITablePartStyle Row { get; } // Band1H

        ITablePartStyle AlternatingRow { get; } // Band2H

        ITablePartStyle Column { get; } // Band1V

        ITablePartStyle AlternatingColumn { get; } // Band2V

        ITablePartStyle LastColumn { get; }

        ITablePartStyle FirstColumn { get; }

        ITablePartStyle LastRow { get; }

        ITablePartStyle FirstRow { get; }
    }
}
