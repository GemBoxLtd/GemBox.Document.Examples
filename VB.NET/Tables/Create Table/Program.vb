Imports GemBox.Document
Imports GemBox.Document.Tables
Imports System.Data
Imports System.Linq

Module Program

    Sub Main()
        Example1()
        Example2()
    End Sub

    Sub Example1()
        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim rowCount As Integer = 10
        Dim columnCount As Integer = 5

        Dim document As New DocumentModel()

        Dim section As New Section(document)
        document.Sections.Add(section)

        ' Create a table with 100% width.
        Dim table As New Table(document)
        table.TableFormat.PreferredWidth = New TableWidth(100, TableWidthUnit.Percentage)
        section.Blocks.Add(table)

        For r As Integer = 0 To rowCount - 1

            ' Create a row and add it to table.
            Dim row As New TableRow(document)
            table.Rows.Add(row)

            For c As Integer = 0 To columnCount - 1

                ' Create a cell and add it to row.
                Dim cell As New TableCell(document)
                row.Cells.Add(cell)

                ' Create a paragraph and add it to cell.
                Dim paragraph As New Paragraph(document, $"Cell ({r + 1},{c + 1})")
                cell.Blocks.Add(paragraph)

            Next

        Next

        document.Save("Simple Table.docx")
    End Sub

    Sub Example2()
        ' If using the Professional version, put your serial key below.
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
