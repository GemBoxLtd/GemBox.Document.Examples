Imports System
Imports System.Data
Imports System.Linq
Imports GemBox.Document

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As DocumentModel = DocumentModel.Load("InvoiceForMailMerge.docx")

        Dim numberOfItems As Integer = 10

        ' Fill items (amount and prices).
        Dim dt As New DataTable()
        dt.TableName = "Item"
        dt.Columns.Add("Date", GetType(DateTime))
        dt.Columns.Add("Hours", GetType(Double))
        dt.Columns.Add("Price", GetType(Double))
        dt.Columns.Add("Total", GetType(Double))

        Dim startDate As DateTime = DateTime.Now.AddDays(-numberOfItems)
        For i As Integer = 0 To numberOfItems - 1
            ' We worked between 6 and 8 hours per day.
            Dim workHours As Integer = i Mod 3 + 6
            Dim totalPrice As Double = workHours * 35

            dt.Rows.Add(startDate.AddDays(i), workHours, 35, totalPrice)
        Next

        document.MailMerge.Execute(dt)

        ' Fill invoice number and date.
        Dim invoiceNumber As Integer = 14
        document.MailMerge.Execute(New With {
            .Number = invoiceNumber,
            .InvoiceDate = DateTime.Now.ToShortDateString()
        })

        ' Fill customer data.
        document.MailMerge.Execute(New With {
            .Name = "ACME Corp",
            .Address = "240 Old Country Road, Springfield, IL",
            .Country = "USA",
            .ContactPerson = "Joe Smith"
        })

        ' Fill total.
        document.MailMerge.Execute(New With {
            .TotalPrice = dt.Rows.Cast(Of DataRow)().Sum(Function(row) CDbl(row("Total")))
        })

        ' Fill notes.
        document.MailMerge.Execute(New With {
            .Notes = "Payment via check."
        })

        document.Save("Merge Ranges.docx")

    End Sub

End Module