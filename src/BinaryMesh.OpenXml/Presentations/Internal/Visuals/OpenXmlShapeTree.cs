using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;

using BinaryMesh.OpenXml.Helpers;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    using Charts = DocumentFormat.OpenXml.Drawing.Charts;

    internal sealed class OpenXmlShapeTree : IShapeTree, IVisualContainer
    {
        private readonly IOpenXmlVisualContainer container;

        private readonly ShapeTree shapeTree;

        public OpenXmlShapeTree(IOpenXmlVisualContainer container, ShapeTree shapeTree)
        {
            this.container = container;
            this.shapeTree = shapeTree;
        }

        public KeyedReadOnlyList<string, IVisual> Visuals => new EnumerableKeyedList<IOpenXmlVisual, string, IVisual>(
            this.shapeTree.Select(element => OpenXmlVisualFactory.TryCreateVisual(this.container, element, out IOpenXmlVisual visual) ? visual : null).Where(visual => visual != null),
            visual => visual.Name,
            visual => visual
        );

        public IShapeVisual PrependShapeVisual(string name)
        {
            throw new NotImplementedException();
        }

        public IShapeVisual InsertShapeVisual(string name, int index)
        {
            throw new NotImplementedException();
        }

        public IShapeVisual AppendShapeVisual(string name)
        {
            Shape shape = this.shapeTree.AppendChild(new Shape()
            {
                NonVisualShapeProperties = new NonVisualShapeProperties()
                {
                    NonVisualDrawingProperties = new NonVisualDrawingProperties()
                    {
                        Name = name,
                        Id = 7 // TODO: calculate
                    },
                    NonVisualShapeDrawingProperties = new NonVisualShapeDrawingProperties(),
                    ApplicationNonVisualDrawingProperties = new ApplicationNonVisualDrawingProperties()
                },
                ShapeProperties = new ShapeProperties()
                    .AppendChildFluent(
                        new Drawing.PresetGeometry() { Preset = Drawing.ShapeTypeValues.Rectangle, AdjustValueList = new Drawing.AdjustValueList() }
                    ),
                TextBody = new TextBody()
                {
                    BodyProperties = new Drawing.BodyProperties(),
                    ListStyle = new Drawing.ListStyle()
                }
            });

            return new OpenXmlShapeVisual(this.container, shape);
        }

        public IConnectionVisual PrependConnectionVisual(string name)
        {
            throw new NotImplementedException();
        }

        public IConnectionVisual InsertConnectionVisual(string name, int index)
        {
            throw new NotImplementedException();
        }

        public IConnectionVisual AppendConnectionVisual(string name)
        {
            ConnectionShape connectionShape = this.shapeTree.AppendChild(new ConnectionShape()
            {
                NonVisualConnectionShapeProperties = new NonVisualConnectionShapeProperties()
                {
                    NonVisualDrawingProperties = new NonVisualDrawingProperties()
                    {
                        Name = name,
                        Id = 7 // TODO: calculate
                    },
                    NonVisualConnectorShapeDrawingProperties = new NonVisualConnectorShapeDrawingProperties(),
                    ApplicationNonVisualDrawingProperties = new ApplicationNonVisualDrawingProperties()
                },
                ShapeProperties = new ShapeProperties()
                    .AppendChildFluent(
                        new Drawing.PresetGeometry() { Preset = Drawing.ShapeTypeValues.Line, AdjustValueList = new Drawing.AdjustValueList() }
                    )
            });

            return new OpenXmlConnectionVisual(this.container, connectionShape);
        }

        public ITableVisual PrependTableVisual(string name)
        {
            throw new NotImplementedException();
        }

        public ITableVisual InsertTableVisual(string name, int index)
        {
            throw new NotImplementedException();
        }

        public ITableVisual AppendTableVisual(string name)
        {
            Drawing.Table table = new Drawing.Table()
                .AppendChildFluent(new Drawing.TableProperties())
                .AppendChildFluent(new Drawing.TableGrid());

            GraphicFrame graphicFrame = this.shapeTree.AppendChild(new GraphicFrame()
            {
                NonVisualGraphicFrameProperties = new NonVisualGraphicFrameProperties()
                {
                    NonVisualDrawingProperties = new NonVisualDrawingProperties()
                    {
                        Name = name,
                        Id = 6 // TODO: calculate
                    },
                    NonVisualGraphicFrameDrawingProperties = new NonVisualGraphicFrameDrawingProperties(),
                    ApplicationNonVisualDrawingProperties = new ApplicationNonVisualDrawingProperties()
                },
                Graphic = new Drawing.Graphic()
                {
                    GraphicData = new Drawing.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/table" }
                        .AppendChildFluent(table)
                }
            });

            return new OpenXmlTableVisual(this.container, graphicFrame);
        }

        public IChartVisual PrependChartVisual(string name)
        {
            throw new NotImplementedException();
        }

        public IChartVisual InsertChartVisual(string name, int index)
        {
            throw new NotImplementedException();
        }

        public IChartVisual AppendChartVisual(string name)
        {
            ChartPart chartPart = this.container.Part.AddNewPartDefaultId<ChartPart>();
            chartPart.ChartSpace = new Charts.ChartSpace() { Date1904 = new Charts.Date1904() { Val = false } }
                .AppendChildFluent(new Charts.Chart());

            Charts.Chart chart = new Charts.Chart();
                chart.SetAttribute(new OpenXmlAttribute("r", "id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", this.container.Part.GetIdOfPart(chartPart)));

            GraphicFrame graphicFrame = this.shapeTree.AppendChild(new GraphicFrame()
            {
                NonVisualGraphicFrameProperties = new NonVisualGraphicFrameProperties()
                {
                    NonVisualDrawingProperties = new NonVisualDrawingProperties()
                    {
                        Name = name,
                        Id = 6 // TODO: calculate
                    },
                    NonVisualGraphicFrameDrawingProperties = new NonVisualGraphicFrameDrawingProperties(),
                    ApplicationNonVisualDrawingProperties = new ApplicationNonVisualDrawingProperties()
                },
                Graphic = new Drawing.Graphic()
                {
                    GraphicData = new Drawing.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" }
                        .AppendChildFluent(chart)
                }
            });

            return new OpenXmlChartVisual(this.container, graphicFrame);
        }
    }
}
