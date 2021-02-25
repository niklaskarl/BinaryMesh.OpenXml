using System;
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
                ISlide slide = presentation.InsertSlide(presentation.SlideMasters[0].SlideLayouts[0]);

                slide.VisualTree["Titel 1"].AsShapeVisual()
                    .SetText("Automated Presentation Documents made easy");

                slide.VisualTree["Untertitel 2"].AsShapeVisual()
                    .SetText("BinaryMesh.OpenXml is an open-source library to easily and intuitively create OpenXml documents");

                slide.VisualTree["Datumsplatzhalter 3"].AsShapeVisual()
                    .SetText("10.10.2020");

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
