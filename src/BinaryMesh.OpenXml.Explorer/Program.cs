using System;
using System.Diagnostics;
using System.IO;
using System.Linq;

using BinaryMesh.OpenXml.Presentations;
using BinaryMesh.OpenXml.Spreadsheets;
using DocumentFormat.OpenXml.Drawing;

namespace BinaryMesh.OpenXml.Explorer
{
    public static class Program
    {
        static void Main(string[] args)
        {
            if (args.Length > 0)
            {
                switch (args[0])
                {
                    case "presentation":
                        GeneratePresentation(args.Skip(1).ToArray());
                        break;
                    case "spreadsheet":
                        GenerateSpreadsheet(args.Skip(1).ToArray());
                        break;
                }
            }
            else
            {
                Console.WriteLine("mode not specified");
            }
        }

        private static void GeneratePresentation(string[] args)
        {
            string destination;
            if (args.Length > 0)
            {
                destination = args[0];
            }
            else
            {
                Console.WriteLine("destination not specified");
                return;
            }

            IPresentation presentation = null;
            using (Stream source = typeof(Program).Assembly.GetManifestResourceStream("BinaryMesh.OpenXml.Explorer.Assets.ExamplePresentation.pptx"))
            {
                presentation = PresentationFactory.CreatePresentation(source);
            }

            using (presentation)
            {
                ISlide titleSlide = presentation.InsertSlide(presentation.SlideMasters[0].SlideLayouts[0]);
                (titleSlide.ShapeTree.Visuals["Titel 1"] as IShapeVisual).Text.SetText("Automated Presentation Documents made easy");
                (titleSlide.ShapeTree.Visuals["Untertitel 2"] as IShapeVisual).Text.SetText("BinaryMesh.OpenXml is an open-source library to easily and intuitively create OpenXml documents");
                (titleSlide.ShapeTree.Visuals["Datumsplatzhalter 3"] as IShapeVisual).Text.SetText("10.10.2020");

                ISlide chartSlide = presentation.InsertSlide(presentation.SlideMasters[0].SlideLayouts[6]);
                IChartVisual pieChartVisual = chartSlide.ShapeTree.AppendChartVisual("Chart 1")
                    .Transform.SetOffset(2032000, 719666)
                    .Transform.SetExtents(8128000, 5418667);

                IChartSpace pieChartSpace = pieChartVisual.ChartSpace;
                using (ISpreadsheetDocument spreadsheet = pieChartSpace.OpenSpreadsheetDocument())
                {
                    IWorkbook workbook = spreadsheet.Workbook;
                    IWorksheet sheet = workbook.AppendWorksheet("Sheet1");

                    string reference = sheet.Cells[0, 1].SetValue("Costs").Reference;

                    sheet.Cells[1, 0].SetValue("1. Quarter");
                    sheet.Cells[2, 0].SetValue("2. Quarter");
                    sheet.Cells[3, 0].SetValue("3. Quarter");
                    sheet.Cells[4, 0].SetValue("4. Quarter");

                    sheet.Cells[1, 1].SetValue(152306);
                    sheet.Cells[2, 1].SetValue(128742);
                    sheet.Cells[3, 1].SetValue(218737);
                    sheet.Cells[4, 1].SetValue(187025);

                    IPieChart pieChart = pieChartSpace.InsertPieChart()
                        .SetFirstSliceAngle(Math.PI * 0.5)
                        .SetExplosion(0.8)
                        .SetHoleSize(0.5);

                    pieChart.Series
                        .SetText(workbook.GetRange("Sheet1!$A$2"))
                        .SetCategoryAxis(workbook.GetRange("Sheet1!$B$1:$E$1"))
                        .SetValueAxis(workbook.GetRange("Sheet1!B$2:$E$2"))
                        .DataLabel.SetShowValue(true)
                        .DataLabel.SetShowSeriesName(false)
                        .DataLabel.SetShowCategoryName(false)
                        .DataLabel.SetShowLegendKey(false)
                        .DataLabel.SetShowPercent(false)
                        .DataLabel.Text.SetFontColor(OpenXmlColor.Rgb(0xFFFFFF));

                    pieChart.Series.Values[0].Style.SetFill(OpenXmlColor.Rgb(0x00FFFF));
                    pieChart.Series.Values[1].Style.SetFill(OpenXmlColor.Rgb(0xFFFFFF));
                    pieChart.Series.Values[2].Style.SetFill(OpenXmlColor.Rgb(0xFFFF00));
                    pieChart.Series.Values[3].Style.SetFill(OpenXmlColor.Rgb(0xFF0000));
                }

                IChartVisual barChartVisual = chartSlide.ShapeTree.AppendChartVisual("Chart 2")
                    .Transform.SetOffset(2032000, 719666)
                    .Transform.SetExtents(8128000, 5418667);

                IChartSpace barChartSpace = barChartVisual.ChartSpace;
                using (ISpreadsheetDocument spreadsheet = barChartSpace.OpenSpreadsheetDocument())
                {
                    IWorkbook workbook = spreadsheet.Workbook;
                    IWorksheet sheet = workbook.AppendWorksheet("Sheet1");

                    string reference = sheet.Cells[0, 1].SetValue("Costs").Reference;

                    sheet.Cells["A2"].SetValue("Kategorie 1");
                    sheet.Cells["A3"].SetValue("Kategorie 2");
                    sheet.Cells["A4"].SetValue("Kategorie 3");
                    sheet.Cells["A5"].SetValue("Kategorie 4");

                    sheet.Cells["B1"].SetValue("Label 1");
                    sheet.Cells["C1"].SetValue("Label 2");
                    sheet.Cells["D1"].SetValue("Label 3");

                    sheet.Cells["B2"].SetValue(106);
                    sheet.Cells["B3"].SetValue(18742);
                    sheet.Cells["B4"].SetValue(237);
                    sheet.Cells["B5"].SetValue(1025);

                    sheet.Cells["C2"].SetValue(12306);
                    sheet.Cells["C3"].SetValue(3441);
                    sheet.Cells["C4"].SetValue(325234);
                    sheet.Cells["C5"].SetValue(123);

                    sheet.Cells["D2"].SetValue(25241);
                    sheet.Cells["D3"].SetValue(8345);
                    sheet.Cells["D4"].SetValue(132523);
                    sheet.Cells["D5"].SetValue(12345);

                    CartesianAxes axes = barChartSpace.AppendCartesianAxes();
                    axes.CategoryAxis.SetVisibility(false);

                    IBarChart barChart = barChartSpace.InsertBarChart(axes)
                        .SetDirection(BarChartDirection.Column)
                        .SetGrouping(BarChartGrouping.Clustered)
                        .InitializeFromRange(sheet.GetRange("B1:D1"), sheet.GetRange("A2:A5"));

                    barChart.Series[0]
                        .DataLabel.SetShowValue(true)
                        .DataLabel.SetShowSeriesName(false)
                        .DataLabel.SetShowCategoryName(false)
                        .DataLabel.SetShowLegendKey(false)
                        .DataLabel.Text.SetFontColor(OpenXmlColor.Rgb(0x00000));

                    barChart.Series[0].Values[0]
                        .DataLabel.SetDelete(true);

                    barChart.Series[0].Values[1]
                        .DataLabel.Style.SetFill(OpenXmlColor.Accent4)
                        .DataLabel.Text.SetFontColor(OpenXmlColor.Rgb(0xFF0000));

                    barChart.Series[1]
                        .DataLabel.SetShowValue(true)
                        .DataLabel.SetShowSeriesName(false)
                        .DataLabel.SetShowCategoryName(false)
                        .DataLabel.SetShowLegendKey(false)
                        .DataLabel.Style.SetFill(OpenXmlColor.Rgb(0x000000))
                        .DataLabel.Text.SetFontColor(OpenXmlColor.Rgb(0xFFFFFF));

                    barChart.Series[0].Values[1].Style.SetFill(OpenXmlColor.Accent2);
                    barChart.Series[1].Values[1].Style.SetFill(OpenXmlColor.Accent2);
                    barChart.Series[2].Values[1].Style.SetFill(OpenXmlColor.Accent2);

                    barChartSpace.CategoryAxes[0].Text.SetFontSize(8);
                    barChartSpace.ValueAxes[0].Text.SetFontSize(8);
                }

                chartSlide.ShapeTree.AppendShapeVisual("Shape 8")
                    .Transform.SetOffset(OpenXmlUnit.Cm(3), OpenXmlUnit.Cm(3))
                    .Transform.SetExtents(OpenXmlUnit.Cm(10), OpenXmlUnit.Cm(10))
                    .Text.SetText("TEST 2")
                    .Text.SetFontSize(8)
                    .Text.SetTextAlign(TextAlignmentTypeValues.Center)
                    .Text.SetTextAnchor(TextAnchoringTypeValues.Center)
                    .Text.SetIsBold(true)
                    .Style.SetFill(OpenXmlColor.Accent4)
                    .Style.SetStroke(OpenXmlColor.Rgb(0, 0, 255))
                    .Style.SetPresetGeometry(OpenXmlPresetGeometry.BuildChevron(28868));

                ISlide tableSlide = presentation.InsertSlide(presentation.SlideMasters[0].SlideLayouts[6]);
                ITableVisual table = tableSlide.ShapeTree.AppendTableVisual("Table 1")
                    .Transform.SetOffset(OpenXmlUnit.Cm(5), OpenXmlUnit.Cm(5));

                table.AppendColumn(OpenXmlUnit.Cm(5));
                table.AppendColumn(OpenXmlUnit.Cm(5));
                table.AppendRow(OpenXmlUnit.Cm(0));
                table.AppendRow(OpenXmlUnit.Cm(0));

                table.Cells[0, 0].Text.SetText("Hello").Text.SetFontSize(8).Text.SetFont("Arial");
                table.Cells[1, 0].Text.SetText("World").Text.SetIsBold(true).Text.SetFont("Comic Sans MS");
                table.Cells[0, 1].Text.SetText("ABC").Text.SetFontColor(OpenXmlColor.Accent2);
                table.Cells[1, 1].Text.SetText("123").Text.SetFontColor(OpenXmlColor.Rgb(25, 240, 120));

                using (Stream stream = new FileStream(destination, FileMode.Create, FileAccess.ReadWrite))
                {
                    presentation.Close(stream);
                }
            }

            new Process
            {
                StartInfo = new ProcessStartInfo(destination)
                {
                    UseShellExecute = true
                }
            }.Start();
        }

        private static void GenerateSpreadsheet(string[] args)
        {
            string destination;
            if (args.Length > 0)
            {
                destination = args[0];
            }
            else
            {
                Console.WriteLine("destination not specified");
                return;
            }

            using (ISpreadsheetDocument document = SpreadsheetFactory.CreateSpreadsheet())
            {
                IWorksheet sheet = document.Workbook.AppendWorksheet("Sheet 1");
                sheet.Cells[0, 0].SetValue("1. Quarter");
                sheet.Cells[1, 0].SetValue("2. Quarter");
                sheet.Cells[2, 0].SetValue("3. Quarter");
                sheet.Cells[3, 0].SetValue("4. Quarter");

                sheet.Cells[0, 1].SetValue(152306);
                sheet.Cells[1, 1].SetValue(128742);
                sheet.Cells[2, 1].SetValue(218737);
                sheet.Cells[3, 1].SetValue(187025);

                sheet.Cells[0, 2].SetValue(90123);
                sheet.Cells[1, 2].SetValue(120744);
                sheet.Cells[2, 2].SetValue(218681);
                sheet.Cells[3, 2].SetValue(187322);

                using (Stream stream = new FileStream(destination, FileMode.Create, FileAccess.ReadWrite))
                {
                    document.Close(stream);
                }
            }

            new Process
            {
                StartInfo = new ProcessStartInfo(destination)
                {
                    UseShellExecute = true
                }
            }.Start();
        }
    }
}
