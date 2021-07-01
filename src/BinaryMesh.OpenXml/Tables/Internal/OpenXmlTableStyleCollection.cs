using System;
using System.Collections.Generic;
using BinaryMesh.OpenXml.Helpers;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;

namespace BinaryMesh.OpenXml.Tables.Internal
{
    internal class OpenXmlTableStyleCollection : BaseEnumerableList<TableStyleEntry, ITableStyle>, ITableStyleCollection
    {
        private readonly Func<bool, TableStylesPart> part;

        public OpenXmlTableStyleCollection(Func<bool, TableStylesPart> part)
        {
            this.part = part;
        }

        protected override IEnumerable<TableStyleEntry> Enumerable => this.GetTableStyleList(false)?.Elements<TableStyleEntry>() ?? new TableStyleEntry[] {};

        protected override Func<TableStyleEntry, ITableStyle> Selector => this.GetTableStyleFromElement;

        public ITableStyle AddTableStyle(string name)
        {
            Guid id = Guid.NewGuid();
            TableStyleList tableStyleList = this.GetTableStyleList(true);
            TableStyleEntry tableStyleEntry = OpenXmlTableStyle.InitializeTableStyle();

            tableStyleEntry.StyleId = id.ToString("B");
            tableStyleEntry.StyleName = name;

            tableStyleList.AppendChild(tableStyleEntry);

            return this.GetTableStyleFromElement(tableStyleEntry);
        }

        private TableStyleList GetTableStyleList(bool create)
        {
            TableStylesPart part = this.part(create);
            TableStyleList list = part?.TableStyleList;
            if (list == null && create)
            {
                list = part.TableStyleList = new TableStyleList();
            }

            return list;
        }

        private ITableStyle GetTableStyleFromElement(TableStyleEntry element)
        {
            return new OpenXmlTableStyle(element);
        }
    }
}