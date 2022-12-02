Imports System.Linq
Imports GemBox.Document
Imports GemBox.Spreadsheet
Imports GemBox.Spreadsheet.Charts

Module Program

    Sub Main()

        ' If using the Professional version, put your GemBox.Document serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        ' If using the Professional version, put your GemBox.Spreadsheet serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Example1()
        Example2()

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

        ' Get underlying Excel chart
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

        ' Add new line series which has doubled values from the first series.
        lineChart.Series.Add("Series 3", lineChart.Series.First() _
            .Values.Cast(Of Double)().Select(Function(val) val * 2))

        document.Save("Updated Chart.docx")
    End Sub

End Module