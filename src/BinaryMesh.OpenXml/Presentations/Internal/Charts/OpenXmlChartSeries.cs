using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using Drawing = DocumentFormat.OpenXml.Drawing;

using BinaryMesh.OpenXml.Spreadsheets;
using System.Linq;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal sealed class OpenXmlChartSeries : IChartSeries, IOpenXmlDataLabelElement
    {
        private readonly OpenXmlElement series;

        public OpenXmlChartSeries(OpenXmlElement series)
        {
            this.series = series;
        }

        public IDataLabel<IChartSeries> DataLabel => new OpenXmlDataLabel<OpenXmlChartSeries, IChartSeries>(this, this);

        public IChartSeries SetText(IRange range)
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

            return this;
        }

        public IChartSeries SetCategoryAxis(IRange range)
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

            return this;
        }

        public IChartSeries SetValueAxis(IRange range)
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

            return this;
        }

        public IChartSeries SetFill(uint index, string srgb)
        {
            DataPoint dataPoint = this.GetOrCreateDataPoint(index);
            this.RemoveFill(dataPoint);

            dataPoint.ChartShapeProperties.AppendChild(
                new Drawing.SolidFill() { RgbColorModelHex = new Drawing.RgbColorModelHex() { Val = srgb } }
            );

            return this;
        }

        private DataPoint GetOrCreateDataPoint(uint index)
        {
            DataPoint dataPoint = this.series.Elements<DataPoint>().FirstOrDefault(dp => dp.Index?.Val == index);
            if (dataPoint == null)
            {
                dataPoint = this.series.AppendChild(
                    new DataPoint()
                    {
                        Index = new Index() { Val = index },
                        Bubble3D = new Bubble3D() { Val = false },
                        ChartShapeProperties = new ChartShapeProperties()
                    }
                );
            }

            return dataPoint;
        }

        public OpenXmlElement GetDataLabel()
        {
            return this.series.GetFirstChild<DataLabels>();
        }

        public OpenXmlElement GetOrCreateDataLabel()
        {
            return this.series.GetFirstChild<DataLabels>() ?? this.series.AppendChild(new DataLabels());
        }

        private void RemoveFill(DataPoint dataPoint)
        {
            dataPoint.ChartShapeProperties.RemoveAllChildren<Drawing.NoFill>();
            dataPoint.ChartShapeProperties.RemoveAllChildren<Drawing.SolidFill>();
            dataPoint.ChartShapeProperties.RemoveAllChildren<Drawing.GradientFill>();
            dataPoint.ChartShapeProperties.RemoveAllChildren<Drawing.BlipFill>();
            dataPoint.ChartShapeProperties.RemoveAllChildren<Drawing.PatternFill>();
            dataPoint.ChartShapeProperties.RemoveAllChildren<Drawing.GroupFill>();
        }
    }
}
