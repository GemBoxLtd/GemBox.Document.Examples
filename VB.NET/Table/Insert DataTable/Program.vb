Imports System.Data
Imports System.Linq
Imports GemBox.Document
Imports GemBox.Document.Tables

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim rowCount As Integer = 10
        Dim columnCount As Integer = 5

        ' Create DataTable with some sample data.
        Dim dataTable As New DataTable()
        For c As Integer = 0 To columnCount - 1
            dataTable.Columns.Add($"Column {c + 1}")
        Next
        For i As Integer = 0 To rowCount - 1
            Dim r = i
            dataTable.Rows.Add(Enumerable.Range(0, columnCount).Select(Function(c) $"Cell ({r + 1},{c + 1})").ToArray())
        Next

        ' Create new document.
        Dim document As New DocumentModel()

        ' Create Table element from DataTable object.
        Dim table As New Table(document, rowCount, columnCount,
            Function(r, c) New TableCell(document, New Paragraph(document, dataTable.Rows(r)(c).ToString())))

        ' Insert first row as Table's header.
        table.Rows.Insert(0, New TableRow(document, dataTable.Columns.Cast(Of DataColumn)().Select(
            Function(dataColumn) New TableCell(document, New Paragraph(document, dataColumn.ColumnName)))))

        table.TableFormat.PreferredWidth = New TableWidth(100, TableWidthUnit.Percentage)

        document.Sections.Add(New Section(document, table))

        document.Save("Insert DataTable.docx")

    End Sub
End Module