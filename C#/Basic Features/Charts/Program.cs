using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using GemBox.Document;
using GemBox.Pdf;
using GemBox.Pdf.Content;
using GemBox.Spreadsheet;
using GemBox.Spreadsheet.Charts;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your GemBox.Document serial key below.
        GemBox.Document.ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        // If using the Professional version, put your GemBox.Spreadsheet serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        // If using the Professional version, put your GemBox.Pdf serial key below.
        GemBox.Pdf.ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        Example1();
        Example2();
        Example3();
        Example4();
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

        // Get underlying Excel sheet and add new cell values.
        var sheet = lineChart.Worksheet;
        sheet.Cells["D1"].Value = "Series 3";
        sheet.Cells["D2"].Value = 8.6;
        sheet.Cells["D3"].Value = 5;
        sheet.Cells["D4"].Value = 7;
        sheet.Cells["D5"].Value = 9;

        // Add new line series to the LineChart.
        lineChart.Series.Add(sheet.Cells["D1"].StringValue, "Sheet1!D2:D5");

        document.Save("Updated Chart.docx");
    }

    static void Example3()
    {
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

    static void Example4()
    {
        var document = DocumentModel.Load("Chart.docx");
        var placeholdersMapping = ReplaceChartsWithPlaceholders(document);
        document.Save("Chart.pdf");

        using (var pdf = PdfDocument.Load("Chart.pdf"))
        {
            ReplacePlaceholdersWithCharts(pdf, placeholdersMapping);
            pdf.Save();
        }
    }

    static readonly string PlaceholderNameFormat = "GemBox_Chart_Placeholder_{0}";
    static readonly Regex PlaceholderNameRegex = new Regex("GemBox_Chart_Placeholder_\\d+");
    static readonly MemoryStream PlaceholderImage = new MemoryStream(File.ReadAllBytes("placeholder.png"));

    static Dictionary<string, MemoryStream> ReplaceChartsWithPlaceholders(DocumentModel document)
    {
        var placeholdersMapping = new Dictionary<string, MemoryStream>();
        int counter = 0;

        foreach (Chart chart in document.GetChildElements(true, ElementType.Chart).Reverse())
        {
            // Replace Word chart with placeholder image that has specific title.
            var placeholder = new Picture(document, PlaceholderImage, PictureFormat.Png, chart.Layout);
            string placeholderName = string.Format(PlaceholderNameFormat, ++counter);
            placeholder.Metadata.Title = placeholderName;
            chart.Content.Start.InsertRange(placeholder.Content);
            chart.Content.Delete();

            // Retrieve Excel chart and export it as PDF.
            var excelChart = (ExcelChart)chart.ExcelChart;
            excelChart.Position.Width = chart.Layout.Size.Width;
            excelChart.Position.Height = chart.Layout.Size.Height;
            var chartAsPdfStream = new MemoryStream();
            excelChart.Format().Save(chartAsPdfStream, GemBox.Spreadsheet.SaveOptions.PdfDefault);

            // Map PDF that contains Excel chart to placeholder name.
            placeholdersMapping.Add(placeholderName, chartAsPdfStream);
        }

        return placeholdersMapping;
    }

    static void ReplacePlaceholdersWithCharts(PdfDocument pdfDocument, Dictionary<string, MemoryStream> placeholdersMapping)
    {
        foreach (var page in pdfDocument.Pages)
        {
            // Find placeholders by searching for images with specific title.
            var placeholders = FindPlaceholders(page);

            foreach (var placeholder in placeholders)
            {
                if (!placeholdersMapping.TryGetValue(placeholder.Key, out MemoryStream chartAsPdfStream))
                    continue;

                PdfImageContent image = placeholder.Value.Item1;
                PdfQuad bounds = placeholder.Value.Item2;

                // Replace placeholder image with PDF that contains Excel chart.
                using (var excelDocument = PdfDocument.Load(chartAsPdfStream))
                {
                    var form = excelDocument.Pages[0].ConvertToForm(pdfDocument);
                    var formContentGroup = page.Content.Elements.AddGroup();
                    var formContent = formContentGroup.Elements.AddForm(form);
                    formContent.Transform = PdfMatrix.CreateTranslation(bounds.Left, bounds.Bottom);
                }

                image.Collection.Remove(image);
            }
        }
    }

    static Dictionary<string, Tuple<PdfImageContent, PdfQuad>> FindPlaceholders(PdfPage page)
    {
        var placeholders = new Dictionary<string, Tuple<PdfImageContent, PdfQuad>>();
        var enumerator = page.Content.Elements.All(page.Transform).GetEnumerator();
        while (enumerator.MoveNext())
        {
            var element = enumerator.Current;
            if (element.ElementType != PdfContentElementType.Image)
                continue;

            var imageElement = (PdfImageContent)element;
            var metadata = imageElement.Image.Metadata?.Value;
            if (metadata == null)
                continue;

            var match = PlaceholderNameRegex.Match(metadata);
            if (!match.Success)
                continue;

            var bounds = imageElement.Bounds;
            enumerator.Transform.Transform(ref bounds);
            placeholders.Add(match.Value, Tuple.Create(imageElement, bounds));
        }

        return placeholders;
    }
}