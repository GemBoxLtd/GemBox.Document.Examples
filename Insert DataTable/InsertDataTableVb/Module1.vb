Imports System
Imports System.Data
Imports GemBox.Document
Imports GemBox.Document.Tables

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As DocumentModel = New DocumentModel

        Dim tableRowCount As Integer = 10
        Dim tableColumnCount As Integer = 5

        ' Initialize DataTable
        Dim dt As New DataTable()

        For i As Integer = 0 To tableColumnCount - 1
            dt.Columns.Add()
        Next

        For i As Integer = 0 To tableRowCount - 1
            Dim row As DataRow = dt.NewRow()
            For j As Integer = 0 To tableColumnCount - 1
                row(j) = String.Format("Cell {0}-{1}", i + 1, j + 1)
            Next
            dt.Rows.Add(row)
        Next

        ' Initialize Table
        Dim table As New Table(document, tableRowCount, tableColumnCount,
            Function(rowIndex As Integer, columnIndex As Integer) New TableCell(document, New Paragraph(document, dt.Rows(rowIndex)(columnIndex).ToString())))

        table.TableFormat.PreferredWidth = New TableWidth(100, TableWidthUnit.Percentage)

        document.Sections.Add(New Section(document, table))

        document.Save("Insert DataTable.docx")

    End Sub

End Module