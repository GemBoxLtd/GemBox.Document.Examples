using GemBox.Document;
using GemBox.Spreadsheet;
using GemBox.Spreadsheet.Charts;
using System.Linq;

class Program
{
    static void Main()
    {
        Example1();
        Example2();
        Example3();
    }

    static void Example1()
    {
        // If using the Professional version, put your GemBox.Document serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        // If using the Professional version, put your GemBox.Spreadsheet serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

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
        // If using the Professional version, put your GemBox.Document serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        // If using the Professional version, put your GemBox.Spreadsheet serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var document = DocumentModel.Load("Chart.docx");

        // Get Word chart.
        var chart = (Chart)document.GetChildElements(true, ElementType.Chart).First();

        // Get underlying Excel chart and cast it as LineChart.
        var lineChart = (LineChart)chart.ExcelChart;

        // Get underlying Excel sheet and add new cell values.
        var sheet = lineChart.Worksheet;
        sheet.Cells["D1"].Value = "Series 3";
        sheet.Cells["D2"].Value = 8.6;
        sheet.Cells["D3"].Value = 5;
        sheet.Cells["D4"].Value = 7;
        sheet.Cells["D5"].Value = 9;

        // Add new line series to the LineChart.
        lineChart.Series.Add(sheet.Cells["D1"].StringValue, "Sheet1!D2:D5");

        lineChart.RefreshCache();

        document.Save("Updated Chart.docx");
    }

    static void Example3()
    {
        // If using the Professional version, put your GemBox.Document serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        // If using the Professional version, put your GemBox.Spreadsheet serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var document = new DocumentModel();

        var chart = new Chart(document, GemBox.Document.ChartType.Column,
            new FloatingLayout(
                new HorizontalPosition(HorizontalPositionType.Center, HorizontalPositionAnchor.Margin),
                new VerticalPosition(VerticalPositionType.Top, VerticalPositionAnchor.Paragraph),
                new Size(10, 5, GemBox.Document.LengthUnit.Centimeter)));

        document.Sections.Add(
            new Section(document,
                new Paragraph(document, chart)));

        // Get underlying Excel chart.
        var columnChart = (ColumnChart)chart.ExcelChart;

        // Set chart's category labels from array.
        columnChart.SetCategoryLabels(new string[] { "Columns 1", "Columns 2", "Columns 3" });

        // Add chart's series from arrays.
        columnChart.Series.Add("Values 1", new double[] { 3.4, 1.1, 3.7 });
        columnChart.Series.Add("Values 2", new double[] { 4.4, 3.9, 3.5 });
        columnChart.Series.Add("Values 3", new double[] { 2.9, 4.1, 1.9 });

        document.Save("Created Chart from Array.docx");
    }
}
