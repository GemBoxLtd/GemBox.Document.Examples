using GemBox.Document;
using GemBox.Spreadsheet;
using GemBox.Spreadsheet.Charts;

class Program
{
    static void Main()
    {
        // If using Professional version, put your GemBox.Document serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        // If using Professional version, put your GemBox.Spreadsheet serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        Example1();
        Example2();
    }

    static void Example1()
    {
        var document = new DocumentModel();

        // Create Word chart and add it to document.
        var chart = new Chart(document, GemBox.Document.ChartType.Bar,
            new FloatingLayout(
                new HorizontalPosition(HorizontalPositionType.Center, HorizontalPositionAnchor.Margin),
                new VerticalPosition(VerticalPositionType.Top, VerticalPositionAnchor.Paragraph),
                new Size(14, 7, GemBox.Document.LengthUnit.Centimeter)));

        document.Sections.Add(
            new Section(document,
                new Paragraph(document, "New document with chart element."),
                new Paragraph(document, chart)));

        // Get underlying Excel chart.
        ExcelChart excelChart = (ExcelChart)chart.ExcelChart;
        ExcelWorksheet worksheet = excelChart.Worksheet;

        // Add data for Excel chart.
        worksheet.Cells["A1"].Value = "Name";
        worksheet.Cells["A2"].Value = "John Doe";
        worksheet.Cells["A3"].Value = "Fred Nurk";
        worksheet.Cells["A4"].Value = "Hans Meier";
        worksheet.Cells["A5"].Value = "Ivan Horvat";

        worksheet.Cells["B1"].Value = "Salary";
        worksheet.Cells["B2"].Value = 3600;
        worksheet.Cells["B3"].Value = 2580;
        worksheet.Cells["B4"].Value = 3200;
        worksheet.Cells["B5"].Value = 4100;

        // Select data.
        excelChart.SelectData(worksheet.Cells.GetSubrange("A1:B5"), true);

        document.Save("Created Chart.pdf");
    }

    static void Example2()
    {
        var document = DocumentModel.Load("Chart.docx");

        // Get Word chart.
        var chart = (Chart)document.GetChildElements(true, ElementType.Chart).First();

        // Get underlying Excel chart and cast it as LineChart.
        var lineChart = (LineChart)chart.ExcelChart;

        // Add new line series which has doubled values from the first series.
        lineChart.Series.Add("Series 3", lineChart.Series.First()
            .Values.Cast<double>().Select(val => val * 2));

        document.Save("Updated Chart.docx");
    }
}