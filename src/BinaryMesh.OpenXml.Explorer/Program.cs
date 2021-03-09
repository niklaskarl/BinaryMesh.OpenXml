﻿using System;
using System.Diagnostics;
using System.IO;
using System.Linq;

using BinaryMesh.OpenXml.Presentations;
using BinaryMesh.OpenXml.Spreadsheets;

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
                titleSlide.ShapeTree.Visuals["Titel 1"].AsShapeVisual().SetText("Automated Presentation Documents made easy");
                titleSlide.ShapeTree.Visuals["Untertitel 2"].AsShapeVisual().SetText("BinaryMesh.OpenXml is an open-source library to easily and intuitively create OpenXml documents");
                titleSlide.ShapeTree.Visuals["Datumsplatzhalter 3"].AsShapeVisual().SetText("10.10.2020");

                ISlide chartSlide = presentation.InsertSlide(presentation.SlideMasters[0].SlideLayouts[6]);
                IChartVisual chartVisual = chartSlide.ShapeTree.AppendChartVisual("Chart 1")
                    .SetOffset(2032000, 719666)
                    .SetExtents(8128000, 5418667);

                IChartSpace chartSpace = chartVisual.ChartSpace;
                using (ISpreadsheetDocument spreadsheet = chartSpace.OpenSpreadsheetDocument())
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

                    IPieChart pieChart = chartSpace.InsertPieChart();
                    pieChart.Series
                        .SetText(workbook.GetRange("Sheet1!$A$2"))
                        .SetCategoryAxis(workbook.GetRange("Sheet1!$B$1:$E$1"))
                        .SetValueAxis(workbook.GetRange("Sheet1!B$2:$E$2"))
                        .SetFill(0, "00FFFF")
                        .SetFill(1, "FFFFFF")
                        .SetFill(2, "FFFF00")
                        .SetFill(3, "FF0000");
                }

                chartSlide.ShapeTree.AppendShapeVisual("Shape 7")
                    .SetOffset(0, 0)
                    .SetExtents(8128000, 5418667)
                    .SetText("TEST")
                    .SetFontSize(8)
                    .SetIsBold(true);

                chartSlide.ShapeTree.AppendShapeVisual("Shape 8")
                    .SetOffset(50000, 0)
                    .SetExtents(8128000, 5418667)
                    .SetText("TEST 2")
                    .SetFontSize(8)
                    .SetIsBold(true)
                    .SetFill(OpenXmlColor.Accent4)
                    .SetStroke(OpenXmlColor.Rgb(0, 0, 255));

                ISlide tableSlide = presentation.InsertSlide(presentation.SlideMasters[0].SlideLayouts[6]);
                ITableVisual table = tableSlide.ShapeTree.AppendTableVisual("Table 1")
                    .SetOffset(OpenXmlUnit.Cm(5), OpenXmlUnit.Cm(5));

                table.AppendColumn(OpenXmlUnit.Cm(5));
                table.AppendColumn(OpenXmlUnit.Cm(5));
                table.AppendRow(OpenXmlUnit.Cm(0));
                table.AppendRow(OpenXmlUnit.Cm(0));

                table.Cells[0, 0].SetText("Hello").SetFontSize(8).SetFont("Arial");
                table.Cells[1, 0].SetText("World").SetIsBold(true).SetFont("Comic Sans MS");
                table.Cells[0, 1].SetText("ABC").SetFontColor(OpenXmlColor.Accent2);
                table.Cells[1, 1].SetText("123").SetFontColor(OpenXmlColor.Rgb(25, 240, 120));

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
