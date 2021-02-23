using System;
using System.IO;

using BinaryMesh.OpenXml.Presentations;

namespace BinaryMesh.OpenXml.Explorer
{
    public static class Program
    {
        static void Main(string[] args)
        {

            IPresentation presentation = null;
            using (Stream source = typeof(Program).Assembly.GetManifestResourceStream("BinaryMesh.OpenXml.Explorer.Assets.ExamplePresentation.pptx"))
            {
                presentation = PresentationFactory.CreatePresentation(source);
            }

            using (presentation)
            {
                ISlide slide = presentation.InsertSlide(presentation.SlideMasters[0].SlideLayouts[0]);

                slide.VisualTree["Titel 1"].AsShapeVisual()
                    .SetText("Statusbericht Versionskontrolle");

                slide.VisualTree["Untertitel 2"].AsShapeVisual()
                    .SetText("VW310 ID.3");

                slide.VisualTree["Datumsplatzhalter 3"].AsShapeVisual()
                    .SetText("22.02.2021");

                using (Stream destination = new FileStream("C:\\Users\\krq\\Desktop\\TestPresentation.pptx", FileMode.Create, FileAccess.ReadWrite))
                {
                    presentation.Close(destination);
                }
            }

            Console.WriteLine("Hello World!");
        }
    }
}
