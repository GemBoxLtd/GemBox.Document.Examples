Imports System
Imports System.Data
Imports System.Linq
Imports GemBox.Document

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim numberOfItems As Integer = 10

        Dim document = DocumentModel.Load("MergeRanges.docx")

        ' Create DataTable as a data source for merge range.
        Dim table As New DataTable() With {.TableName = "Items"}

        table.Columns.Add("Date", GetType(DateTime))
        table.Columns.Add("Hours", GetType(Integer))
        table.Columns.Add("Unit", GetType(Double))
        table.Columns.Add("Price", GetType(Double))

        For rowIndex As Integer = 1 To numberOfItems

            Dim [date] As DateTime = DateTime.Today.AddDays(rowIndex - numberOfItems)
            Dim hours As Integer = rowIndex Mod 3 + 6
            Dim unit As Double = 35.0
            Dim price As Double = hours * unit

            table.Rows.Add([date], hours, unit, price)

        Next

        ' Execute mail merge process for "Items" merge range.
        document.MailMerge.Execute(table)

        ' Execute mail merge process again for "Number", "Date" and "TotalPrice" fields.
        document.MailMerge.Execute(
            New With
            {
                .Number = 10203,
                .Date = DateTime.Now,
                .TotalPrice = table.Rows.Cast(Of DataRow)().Sum(Function(row) CDbl(row("Price")))
            })

        document.Save("Merged Ranges Output.docx")

    End Sub
End Module