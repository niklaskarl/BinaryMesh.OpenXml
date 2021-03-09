# BinaryMesh.OpenXml

🚧🚧🚧 This library is under construction 🚧🚧🚧

## Currently Available Features
### Presentations
 - Creating and opening presentations
 - Inserting slides based on layouts
 - Creating Shapes
 - Editing text of shapes
 - Formating Shapes
 - Creating Charts
### Spreadsheets
 - Creating and opening spreadsheets
 - Inserting sheets
 - Setting Values of Cells
 - Reading Cells

## Planed Features
### Presentations
 - Advanced Shapes
 - Pictures
 - Even more formatting
 - Even more charts
 - Rich Text in Shapes
 - ...
### Spreadsheets
 - Formatting
 - ...
### Wordprocessing
 - Basic support
 - ...

## Example API for creating Presentations

``` csharp
using BinaryMesh.OpenXml.Presentations;
using BinaryMesh.OpenXml.Spreadsheets;

using (IPresentationDocument document = PresentationFactory.CreatePresentationDocument())
{
    IPresentation = document.Presentation;
    ISlide titleSlide = presentation.InsertSlide(presentation.SlideMasters["Office"].SlideLayouts["Title"]);
    titleSlide.VisualTree["Title 1"].AsShapeVisual()
        .SetText("Automated Presentation Documents made easy")
    titleSlide.VisualTree["Subtitle 2"].AsShapeVisual()
        .SetText("BinaryMesh.OpenXml is an open-source library to easily and intuitively create OpenXml documents")

    IChartSpace pieChartSpace = presentation.CreateChartSpace();
    IPieChart pieChart = pieChartSpace.InsertPieChart();
    using (ISpreadsheetDocument spreadsheet = pieChartSpace.OpenSpreadsheetDocument())
    {
        ISheet sheet = spreadsheet.Workbook.Sheets["Sheet 1"];
        sheet.Cells["B1"].SetValue("Revenue");
        sheet.Cells["A2"].SetValue("2019");
        sheet.Cells["A2"].SetValue(1000);
        sheet.Cells["B3"].SetValue("2020");
        sheet.Cells["B3"].SetValue(1875.13);
        sheet.Cells["C4"].SetValue("2021");
        sheet.Cells["C4"].SetValue(874.86);

        pieChart.Series.SetTextRange(sheet.GetRange("$B$1"));
        pieChart.Series.SetCategoryRange(sheet.GetRange("$A$2:$A$5"));
        pieChart.Series.SetValueRange(sheet.GetRange("$B$2:$B$5"));
    }

    ISlide statisticsSlide = presentation.InsertSlide(presentation.SlideMasters["Office"].SlideLayouts["Only Title"]);
    titleSlide.VisualTree["Title 1"].AsShapeVisual()
        .SetText("A basic Chart");
    titleSlide.VisualTree.InsertGraphicFrame("Diagramm 1")
        .SetOffset(2032000, 719666)
        .SetExtend(8128000, 5418667)
        .SetContent(pieChartSpace);

    // write presentation to a destination file and close it.
    using (Stream destination = new FileStream("presentation.pptx"))
    {
        document.Close(destination);
    }
}

```