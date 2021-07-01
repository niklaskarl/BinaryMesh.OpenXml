using System;
using System.Collections.Generic;

namespace BinaryMesh.OpenXml.Tables
{
    public interface ITableStyleCollection : IReadOnlyList<ITableStyle>
    {
        ITableStyle AddTableStyle(string name);
    }
}
