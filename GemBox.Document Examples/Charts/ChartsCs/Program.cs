using GemBox.Document;
using GemBox.Spreadsheet;
using GemBox.Spreadsheet.Charts;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        // This is needed to enable chart rendering in the GemBox.Document.
        // If you own a professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        Example1();
        Example2();
    }

    static void Example1()
    {
        // Load input file and save it in selected output format.
        DocumentModel document = DocumentModel.Load("Charts.docx");
        document.Save("Output.pdf");
    }

    static void Example2()
    {
        DocumentModel document = new DocumentModel();

        // Create simple document title.
        Paragraph title = new Paragraph(document, "GemBox.Document - Create chart example");

        title.ParagraphFormat.Alignment = HorizontalAlignment.Center;

        // Create chart.
        Chart chart = new Chart(document, GemBox.Document.ChartType.Bar,
            new FloatingLayout(
                new HorizontalPosition(HorizontalPositionType.Center, HorizontalPositionAnchor.Margin),
                new VerticalPosition(VerticalPositionType.Top, VerticalPositionAnchor.Paragraph),
                new Size(14, 7, GemBox.Document.LengthUnit.Centimeter)));

        // Get underlying Excel chart.
        ExcelChart excelChart = (ExcelChart)chart.ExcelChart;
        ExcelWorksheet worksheet = excelChart.Worksheet;

        // Add data which will be used by the Excel chart.
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
        excelChart.SelectData(
            worksheet.Cells.GetSubrangeAbsolute(0, 0, 4, 1), true);

        // Add document elements.
        document.Sections.Add(new Section(document,
            title, new Paragraph(document, chart)));

        document.Save("Chart.docx");
    }
}