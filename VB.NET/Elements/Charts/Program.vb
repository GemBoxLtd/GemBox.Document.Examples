Imports GemBox.Document
Imports GemBox.Spreadsheet
Imports GemBox.Spreadsheet.Charts
Imports System.Linq

Module Program

    Sub Main()
        Example1()
        Example2()
        Example3()
    End Sub

    Sub Example1()
        ' If using the Professional version, put your GemBox.Document serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        ' If using the Professional version, put your GemBox.Spreadsheet serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

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
        ' If using the Professional version, put your GemBox.Document serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        ' If using the Professional version, put your GemBox.Spreadsheet serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

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
        ' If using the Professional version, put your GemBox.Document serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        ' If using the Professional version, put your GemBox.Spreadsheet serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

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

End Module
