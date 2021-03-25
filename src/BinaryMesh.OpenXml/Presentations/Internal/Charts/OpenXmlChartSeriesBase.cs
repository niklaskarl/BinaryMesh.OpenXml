using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using Drawing = DocumentFormat.OpenXml.Drawing;

using BinaryMesh.OpenXml.Spreadsheets;
using System.Linq;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal abstract class OpenXmlChartSeriesBase<TSelf>
    {
        protected readonly OpenXmlElement series;

        public OpenXmlChartSeriesBase(OpenXmlElement series)
        {
            this.series = series;
        }

        protected abstract TSelf Self { get; }

        public TSelf SetText(IRange range)
        {
            StringCache stringCache = new StringCache();
            if (range.Width == 1 && range.Height > 0)
            {
                stringCache.PointCount = new PointCount() { Val = (uint)range.Height.Value };
                for (uint i = 0; i < range.Height; ++i)
                {
                    stringCache.AppendChild(new StringPoint() { Index = i, NumericValue = new NumericValue(range[0, i].InnerValue) });
                }
            }
            else if (range.Height == 1 && range.Width > 0)
            {
                stringCache.PointCount = new PointCount() { Val = (uint)range.Width.Value };
                for (uint i = 0; i < range.Width; ++i)
                {
                    stringCache.AppendChild(new StringPoint() { Index = i, NumericValue = new NumericValue(range[i, 0].InnerValue) });
                }
            }
            else
            {
                throw new ArgumentException("Expected an one-dimensional range.");
            }

            SeriesText seriesText = series.GetFirstChild<SeriesText>() ?? series.AppendChild(new SeriesText());
            seriesText.StringReference = new StringReference()
            {
                Formula = new Formula(range.Formula),
                StringCache = stringCache
            };

            return this.Self;
        }

        public TSelf SetCategoryAxis(IRange range)
        {
            StringCache stringCache = new StringCache();
            if (range.Width == 1 && range.Height > 0)
            {
                stringCache.PointCount = new PointCount() { Val = (uint)range.Height.Value };
                for (uint i = 0; i < range.Height; ++i)
                {
                    stringCache.AppendChild(new StringPoint() { Index = i, NumericValue = new NumericValue(range[0, i].InnerValue) });
                }
            }
            else if (range.Height == 1 && range.Width > 0)
            {
                stringCache.PointCount = new PointCount() { Val = (uint)range.Width.Value };
                for (uint i = 0; i < range.Width; ++i)
                {
                    stringCache.AppendChild(new StringPoint() { Index = i, NumericValue = new NumericValue(range[i, 0].InnerValue) });
                }
            }
            else
            {
                throw new ArgumentException("Expected an one-dimensional range.");
            }

            CategoryAxisData categoryAxisData = series.GetFirstChild<CategoryAxisData>() ?? series.AppendChild(new CategoryAxisData());
            categoryAxisData.StringReference = new StringReference()
            {
                Formula = new Formula(range.Formula),
                StringCache = stringCache
            };

            return this.Self;
        }

        public TSelf SetValueAxis(IRange range)
        {
            NumberingCache numberingCache = new NumberingCache();
            if (range.Width == 1 && range.Height > 0)
            {
                numberingCache.PointCount = new PointCount() { Val = (uint)range.Height.Value };
                for (uint i = 0; i < range.Height; ++i)
                {
                    numberingCache.AppendChild(new NumericPoint() { Index = i, NumericValue = new NumericValue(range[0, i].InnerValue) });
                }
            }
            else if (range.Height == 1 && range.Width > 0)
            {
                numberingCache.PointCount = new PointCount() { Val = (uint)range.Width.Value };
                for (uint i = 0; i < range.Width; ++i)
                {
                    numberingCache.AppendChild(new NumericPoint() { Index = i, NumericValue = new NumericValue(range[i, 0].InnerValue) });
                }
            }
            else
            {
                throw new ArgumentException("Expected an one-dimensional range.");
            }

            Values values = series.GetFirstChild<Values>() ?? series.AppendChild(new Values());
            values.NumberReference = new NumberReference()
            {
                Formula = new Formula(range.Formula),
                NumberingCache = numberingCache
            };

            return this.Self;
        }
    }
}
