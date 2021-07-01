using System;
using DocumentFormat.OpenXml.Drawing;

using BinaryMesh.OpenXml.Shared;
using BinaryMesh.OpenXml.Styles;
using BinaryMesh.OpenXml.Styles.Internal;

namespace BinaryMesh.OpenXml.Tables.Internal
{
    internal class OpenXmlTableCellStyle : ITableCellStyle
    {
        private readonly OpenXmlTablePartStyle parent;

        private readonly ElementGenerator<TableCellStyle> tableCellStyleGenerator;

        public OpenXmlTableCellStyle(OpenXmlTablePartStyle parent, ElementGenerator<TableCellStyle> tableCellStyleGenerator)
        {
            this.parent = parent;
            this.tableCellStyleGenerator = tableCellStyleGenerator;
        }

        public IFillStyle<ITablePartStyle> Fill => new OpenXmlFillStyle<ITablePartStyle>(this.parent, this.GetFill);

        public ITableCellBoderStyle Border => new OpenXmlTableCellBorderStyle(this);

        private Fill GetFill(bool create)
        {
            TableCellStyle style = this.tableCellStyleGenerator(create);
            Fill result = style?.GetFirstChild<Fill>();
            if (result == null && create)
            {
                result = style.AppendChild(new Fill());
            }

            return result;
        }

        private class OpenXmlTableCellBorderStyle : ITableCellBoderStyle
        {
            private readonly OpenXmlTableCellStyle parent;

            public OpenXmlTableCellBorderStyle(OpenXmlTableCellStyle parent)
            {
                this.parent = parent;
            }

            public IStrokeStyle<ITablePartStyle> Left => new OpenXmlStrokeStyle<ITablePartStyle>(this.parent.parent, this.GetLeftBorderOutline);

            public IStrokeStyle<ITablePartStyle> Top => new OpenXmlStrokeStyle<ITablePartStyle>(this.parent.parent, this.GetTopBorderOutline);

            public IStrokeStyle<ITablePartStyle> Right => new OpenXmlStrokeStyle<ITablePartStyle>(this.parent.parent, this.GetRightBorderOutline);

            public IStrokeStyle<ITablePartStyle> Bottom => new OpenXmlStrokeStyle<ITablePartStyle>(this.parent.parent, this.GetBottomBorderOutline);

            public IStrokeStyle<ITablePartStyle> InsideHorizontal => new OpenXmlStrokeStyle<ITablePartStyle>(this.parent.parent, this.GetInsideHorizontalBorderOutline);

            public IStrokeStyle<ITablePartStyle> InsideVertical => new OpenXmlStrokeStyle<ITablePartStyle>(this.parent.parent, this.GetInsideVerticalBorderOutline);

            private TableCellBorders GetTableCellBorders(bool create)
            {
                TableCellStyle style = this.parent.tableCellStyleGenerator(create);
                TableCellBorders result = style?.TableCellBorders;
                if (result == null && create)
                {
                    result = new TableCellBorders();
                    style.TableCellBorders = result;
                }

                return result;
            }

            private Outline GetLeftBorderOutline(bool create)
            {
                TableCellBorders tableCellBorders = this.GetTableCellBorders(create);
                LeftBorder border = tableCellBorders?.LeftBorder;
                if (border == null && create)
                {
                    border = new LeftBorder();
                    tableCellBorders.LeftBorder = border;
                }

                Outline result = border?.Outline;
                if (result == null && create)
                {
                    result = new Outline();
                    border.Outline = result;
                }

                return result;
            }

            private Outline GetTopBorderOutline(bool create)
            {
                TableCellBorders tableCellBorders = this.GetTableCellBorders(create);
                TopBorder border = tableCellBorders?.TopBorder;
                if (border == null && create)
                {
                    border = new TopBorder();
                    tableCellBorders.TopBorder = border;
                }

                Outline result = border?.Outline;
                if (result == null && create)
                {
                    result = new Outline();
                    border.Outline = result;
                }

                return result;
            }

            private Outline GetRightBorderOutline(bool create)
            {
                TableCellBorders tableCellBorders = this.GetTableCellBorders(create);
                RightBorder border = tableCellBorders?.RightBorder;
                if (border == null && create)
                {
                    border = new RightBorder();
                    tableCellBorders.RightBorder = border;
                }

                Outline result = border?.Outline;
                if (result == null && create)
                {
                    result = new Outline();
                    border.Outline = result;
                }

                return result;
            }

            private Outline GetBottomBorderOutline(bool create)
            {
                TableCellBorders tableCellBorders = this.GetTableCellBorders(create);
                BottomBorder border = tableCellBorders?.BottomBorder;
                if (border == null && create)
                {
                    border = new BottomBorder();
                    tableCellBorders.BottomBorder = border;
                }

                Outline result = border?.Outline;
                if (result == null && create)
                {
                    result = new Outline();
                    border.Outline = result;
                }

                return result;
            }

            private Outline GetInsideHorizontalBorderOutline(bool create)
            {
                TableCellBorders tableCellBorders = this.GetTableCellBorders(create);
                InsideHorizontalBorder border = tableCellBorders?.InsideHorizontalBorder;
                if (border == null && create)
                {
                    border = new InsideHorizontalBorder();
                    tableCellBorders.InsideHorizontalBorder = border;
                }

                Outline result = border?.Outline;
                if (result == null && create)
                {
                    result = new Outline();
                    border.Outline = result;
                }

                return result;
            }

            private Outline GetInsideVerticalBorderOutline(bool create)
            {
                TableCellBorders tableCellBorders = this.GetTableCellBorders(create);
                InsideVerticalBorder border = tableCellBorders?.InsideVerticalBorder;
                if (border == null && create)
                {
                    border = new InsideVerticalBorder();
                    tableCellBorders.InsideVerticalBorder = border;
                }

                Outline result = border?.Outline;
                if (result == null && create)
                {
                    result = new Outline();
                    border.Outline = result;
                }

                return result;
            }

        }
    }
}
