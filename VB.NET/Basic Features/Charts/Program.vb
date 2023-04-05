Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Text.RegularExpressions
Imports GemBox.Document
Imports GemBox.Pdf
Imports GemBox.Pdf.Content
Imports GemBox.Spreadsheet
Imports GemBox.Spreadsheet.Charts

Module Program

    Sub Main()

        ' If using the Professional version, put your GemBox.Document serial key below.
        GemBox.Document.ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        ' If using the Professional version, put your GemBox.Spreadsheet serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        ' If using the Professional version, put your GemBox.Pdf serial key below.
        GemBox.Pdf.ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Example1()
        Example2()
        Example3()
        Example4()

    End Sub

    Sub Example1()
        Dim document As New DocumentModel()

        ' Create Word chart and add it to document.
        Dim chart As New Chart(document, GemBox.Document.ChartType.Bar,
            New FloatingLayout(
                New HorizontalPosition(HorizontalPositionType.Center, HorizontalPositionAnchor.Margin),
                New VerticalPosition(VerticalPositionType.Top, VerticalPositionAnchor.Paragraph),
                New Size(14, 7, GemBox.Document.LengthUnit.Centimeter)))

        document.Sections.Add(
            New Section(document,
                New Paragraph(document, "New document with chart element."),
                New Paragraph(document, chart)))

        ' Get underlying Excel chart.
        Dim excelChart As ExcelChart = DirectCast(chart.ExcelChart, ExcelChart)
        Dim worksheet As ExcelWorksheet = excelChart.Worksheet

        ' Add data for Excel chart.
        worksheet.Cells("A1").Value = "Name"
        worksheet.Cells("A2").Value = "John Doe"
        worksheet.Cells("A3").Value = "Fred Nurk"
        worksheet.Cells("A4").Value = "Hans Meier"
        worksheet.Cells("A5").Value = "Ivan Horvat"

        worksheet.Cells("B1").Value = "Salary"
        worksheet.Cells("B2").Value = 3600
        worksheet.Cells("B3").Value = 2580
        worksheet.Cells("B4").Value = 3200
        worksheet.Cells("B5").Value = 4100

        ' Select data.
        excelChart.SelectData(worksheet.Cells.GetSubrange("A1:B5"), True)

        document.Save("Created Chart.pdf")
    End Sub

    Sub Example2()
        Dim document = DocumentModel.Load("Chart.docx")

        ' Get Word chart.
        Dim chart = DirectCast(document.GetChildElements(True, ElementType.Chart).First(), Chart)

        ' Get underlying Excel chart and cast it as LineChart.
        Dim lineChart = DirectCast(chart.ExcelChart, LineChart)

        ' Get underlying Excel sheet and add new cell values.
        Dim sheet = lineChart.Worksheet
        sheet.Cells("D1").Value = "Series 3"
        sheet.Cells("D2").Value = 8.6
        sheet.Cells("D3").Value = 5
        sheet.Cells("D4").Value = 7
        sheet.Cells("D5").Value = 9

        ' Add new line series to the LineChart.
        lineChart.Series.Add(sheet.Cells("D1").StringValue, "Sheet1!D2:D5")

        document.Save("Updated Chart.docx")
    End Sub

    Sub Example3()
        Dim document As New DocumentModel()

        Dim chart As New Chart(document, GemBox.Document.ChartType.Column,
            New FloatingLayout(
                New HorizontalPosition(HorizontalPositionType.Center, HorizontalPositionAnchor.Margin),
                New VerticalPosition(VerticalPositionType.Top, VerticalPositionAnchor.Paragraph),
                New Size(10, 5, GemBox.Document.LengthUnit.Centimeter)))

        document.Sections.Add(
            New Section(document,
                New Paragraph(document, chart)))

        ' Get underlying Excel chart.
        Dim columnChart = DirectCast(chart.ExcelChart, ColumnChart)

        ' Set chart's category labels from array.
        columnChart.SetCategoryLabels(New String() {"Columns 1", "Columns 2", "Columns 3"})

        ' Add chart's series from arrays.
        columnChart.Series.Add("Values 1", New Double() {3.4, 1.1, 3.7})
        columnChart.Series.Add("Values 2", New Double() {4.4, 3.9, 3.5})
        columnChart.Series.Add("Values 3", New Double() {2.9, 4.1, 1.9})

        document.Save("Created Chart from Array.docx")
    End Sub

    Sub Example4()
        Dim document = DocumentModel.Load("Chart.docx")
        Dim placeholdersMapping = ReplaceChartsWithPlaceholders(document)
        document.Save("Chart.pdf")

        Using pdf = PdfDocument.Load("Chart.pdf")
            ReplacePlaceholdersWithCharts(pdf, placeholdersMapping)
            pdf.Save()
        End Using
    End Sub

    ReadOnly PlaceholderNameFormat As String = "GemBox_Chart_Placeholder_{0}"
    ReadOnly PlaceholderNameRegex As Regex = New Regex("GemBox_Chart_Placeholder_\d+")
    ReadOnly PlaceholderImage As MemoryStream = New MemoryStream(File.ReadAllBytes("placeholder.png"))

    Function ReplaceChartsWithPlaceholders(document As DocumentModel) As Dictionary(Of String, MemoryStream)
        Dim placeholdersMapping = New Dictionary(Of String, MemoryStream)()
        Dim counter As Integer = 0

        For Each chart As Chart In document.GetChildElements(True, ElementType.Chart).Reverse()
            ' Replace Word chart with placeholder image that has specific title.
            Dim placeholder = New Picture(document, PlaceholderImage, PictureFormat.Png, chart.Layout)
            counter += 1
            Dim placeholderName As String = String.Format(PlaceholderNameFormat, counter)
            placeholder.Metadata.Title = placeholderName
            chart.Content.Start.InsertRange(placeholder.Content)
            chart.Content.Delete()

            ' Retrieve Excel chart and export it as PDF.
            Dim excelChart = CType(chart.ExcelChart, ExcelChart)
            excelChart.Position.Width = chart.Layout.Size.Width
            excelChart.Position.Height = chart.Layout.Size.Height
            Dim chartAsPdfStream = New MemoryStream()
            excelChart.Format().Save(chartAsPdfStream, GemBox.Spreadsheet.SaveOptions.PdfDefault)

            ' Map PDF that contains Excel chart to placeholder name.
            placeholdersMapping.Add(placeholderName, chartAsPdfStream)
        Next

        Return placeholdersMapping
    End Function

    Sub ReplacePlaceholdersWithCharts(pdfDocument As PdfDocument, placeholdersMapping As Dictionary(Of String, MemoryStream))
        Dim chartAsPdfStream As MemoryStream = Nothing

        For Each page In pdfDocument.Pages
            ' Find placeholders by searching for images with specific title.
            Dim placeholders = FindPlaceholders(page)

            For Each placeholder In placeholders
                If Not placeholdersMapping.TryGetValue(placeholder.Key, chartAsPdfStream) Then Continue For

                Dim image As PdfImageContent = placeholder.Value.Item1
                Dim bounds As PdfQuad = placeholder.Value.Item2

                ' Replace placeholder image with PDF that contains Excel chart.
                Using excelDocument = PdfDocument.Load(chartAsPdfStream)
                    Dim form = excelDocument.Pages(0).ConvertToForm(pdfDocument)
                    Dim formContentGroup = page.Content.Elements.AddGroup()
                    Dim formContent = formContentGroup.Elements.AddForm(form)
                    formContent.Transform = PdfMatrix.CreateTranslation(bounds.Left, bounds.Bottom)
                End Using

                image.Collection.Remove(image)
            Next
        Next
    End Sub

    Function FindPlaceholders(page As PdfPage) As Dictionary(Of String, Tuple(Of PdfImageContent, PdfQuad))
        Dim placeholders = New Dictionary(Of String, Tuple(Of PdfImageContent, PdfQuad))()
        Dim enumerator = page.Content.Elements.All(page.Transform).GetEnumerator()

        While enumerator.MoveNext()

            Dim element = enumerator.Current
            If element.ElementType <> PdfContentElementType.Image Then Continue While

            Dim imageElement = CType(element, PdfImageContent)
            Dim metadata = imageElement.Image.Metadata?.Value
            If metadata Is Nothing Then Continue While

            Dim match = PlaceholderNameRegex.Match(metadata)
            If Not match.Success Then Continue While

            Dim bounds = imageElement.Bounds
            enumerator.Transform.Transform(bounds)
            placeholders.Add(match.Value, Tuple.Create(imageElement, bounds))

        End While

        Return placeholders
    End Function
End Module