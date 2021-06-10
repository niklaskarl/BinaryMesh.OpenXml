using System;
using BinaryMesh.OpenXml.Spreadsheets;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;

namespace BinaryMesh.OpenXml.Charts.Internal
{
    internal sealed class OpenXmlPieChartSeries : OpenXmlChartSeries<IPieChartSeries, IPieChartValue>, IPieChartSeries
    {
        public OpenXmlPieChartSeries(OpenXmlElement series) :
            base(series)
        {
        }

        protected override IPieChartSeries Result => this;

        protected override IPieChartValue ConstructValue(uint index)
        {
            return new OpenXmlPieChartValue(this, index);
        }

        public IPieChartSeries SetText(IRange range)
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

            SeriesText seriesText = this.series.GetFirstChild<SeriesText>() ?? this.series.AppendChild(new SeriesText());
            seriesText.StringReference = new StringReference()
            {
                Formula = new Formula(range.Formula),
                StringCache = stringCache
            };

            return this;
        }

        public IPieChartSeries SetCategoryAxis(IRange range)
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

            CategoryAxisData categoryAxisData = this.series.GetFirstChild<CategoryAxisData>() ?? this.series.AppendChild(new CategoryAxisData());
            categoryAxisData.StringReference = new StringReference()
            {
                Formula = new Formula(range.Formula),
                StringCache = stringCache
            };

            return this;
        }

        public IPieChartSeries SetValueAxis(IRange range)
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

            Values values = this.series.GetFirstChild<Values>() ?? this.series.AppendChild(new Values());
            values.NumberReference = new NumberReference()
            {
                Formula = new Formula(range.Formula),
                NumberingCache = numberingCache
            };

            return this;
        }
    }
}
