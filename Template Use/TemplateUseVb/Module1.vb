Imports System
Imports System.Linq
Imports GemBox.Document
Imports GemBox.Document.Tables

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As DocumentModel = DocumentModel.Load("Invoice.docx")

        Dim numberOfItems As Integer = 10

        ' Document contains 4 tables. Each table contains some set of information.
        Dim tables As Table() = document.GetChildElements(True, ElementType.Table).Cast(Of Table)().ToArray()

        ' First table contains invoice number and date.
        Dim invoiceTable As Table = tables(0)
        Dim invoiceNumber As Integer = 14
        ' We can get and cast first paragraph using:
        ' a) Linq
        invoiceTable.Rows(0).Cells(1).Blocks.Cast(Of Paragraph)().First().Inlines.Add(New Run(document, invoiceNumber.ToString()))
        ' b) ElementCollection.Cast<TElement>(int index)
        invoiceTable.Rows(1).Cells(1).Blocks.Cast(Of Paragraph)(0).Inlines.Add(New Run(document, DateTime.Now.ToShortDateString()))

        ' Second table contains customer data.
        Dim customerTable As Table = tables(1)
        customerTable.Rows(0).Cells(1).Blocks.Add(New Paragraph(document, "ACME Corp"))
        customerTable.Rows(1).Cells(1).Blocks.Add(New Paragraph(document, "240 Old Country Road, Springfield, IL"))
        customerTable.Rows(2).Cells(1).Blocks.Add(New Paragraph(document, "USA"))
        customerTable.Rows(3).Cells(1).Blocks.Add(New Paragraph(document, "Joe Smith"))

        ' Third table contains amount and prices.
        Dim mainTable As Table = tables(2)
        ' Our main table contains only one row (for one item). If we have more items then we need to clone it.
        For i As Integer = 1 To numberOfItems - 1
            mainTable.Rows.Insert(1, mainTable.Rows(1).Clone(True))
        Next

        Dim startDate As DateTime = DateTime.Now.AddDays(-numberOfItems)
        Dim total As Integer = 0

        Dim rowIndex As Integer
        For rowIndex = 1 To numberOfItems
            mainTable.Rows(rowIndex).Cells(0).Blocks.Add(New Paragraph(document, startDate.AddDays(rowIndex - 1).ToString("dddd, MMMM d, yyyy")))
            ' We worked between 6 and 8 hours per day.
            Dim workHours As Integer = rowIndex Mod 3 + 6
            Dim price As Integer = workHours * 35
            total += price
            mainTable.Rows(rowIndex).Cells(1).Blocks.Cast(Of Paragraph)(0).Inlines.Add(New Run(document, workHours.ToString()))
            mainTable.Rows(rowIndex).Cells(2).Blocks.Cast(Of Paragraph)(0).Inlines.Add(New Run(document, "35.00"))
            mainTable.Rows(rowIndex).Cells(3).Blocks.Cast(Of Paragraph)(0).Inlines.Add(New Run(document, price.ToString("0.00")))
        Next

        mainTable.Rows(rowIndex).Cells(3).Blocks.Cast(Of Paragraph)(0).Inlines.Add(New Run(document, total.ToString("0.00")))

        ' Fourth table contains notes
        Dim notesTable As Table = tables(3)
        notesTable.Rows(1).Cells(0).Blocks.Add(New Paragraph(document, "Payment via check."))

        document.Save("Template Use.docx")

    End Sub

End Module