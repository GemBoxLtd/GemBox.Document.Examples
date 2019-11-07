Imports GemBox.Document
Imports GemBox.Spreadsheet
Imports GemBox.Spreadsheet.Charts

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        ' This is needed to enable chart rendering in the GemBox.Document.
        ' If you own a professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Example1()
        Example2()

    End Sub

    Sub Example1
        ' Load input file and save it in selected output format
        Dim document As DocumentModel = DocumentModel.Load("Charts.docx")
        document.Save("Output.pdf")
    End Sub

    Sub Example2
        Dim document As DocumentModel = New DocumentModel()

        ' Create simple document title
        Dim title As Paragraph = New Paragraph(document, "GemBox.Document - Create chart example")

        title.ParagraphFormat.Alignment = HorizontalAlignment.Center

        ' Create chart
        Dim chart As Chart = New Chart(document, GemBox.Document.ChartType.Bar,
            New FloatingLayout(
                New HorizontalPosition(HorizontalPositionType.Center, HorizontalPositionAnchor.Margin),
                New VerticalPosition(VerticalPositionType.Top, VerticalPositionAnchor.Paragraph),
                New Size(14, 7, GemBox.Document.LengthUnit.Centimeter)))

        ' Get underlying Excel chart
        Dim excelChart As ExcelChart = DirectCast(chart.ExcelChart, ExcelChart)
        Dim worksheet As ExcelWorksheet = excelChart.Worksheet

        ' Add data which will be used by the Excel chart.
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

        ' Select data
        excelChart.SelectData(
            worksheet.Cells.GetSubrangeAbsolute(0, 0, 4, 1), True)

        ' Add document elements
        document.Sections.Add(New Section(document,
            title, New Paragraph(document, chart)))

        document.Save("Chart.docx")
    End Sub

End Module