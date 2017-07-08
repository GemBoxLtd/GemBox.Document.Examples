Imports System
Imports GemBox.Document
Imports GemBox.Document.Tables

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As DocumentModel = New DocumentModel

        Dim tableRowCount As Integer = 10
        Dim tableColumnCount As Integer = 5

        Dim table As New Table(document)
        table.TableFormat.PreferredWidth = New TableWidth(100, TableWidthUnit.Percentage)

        For i As Integer = 0 To tableRowCount - 1
            Dim row As New TableRow(document)
            table.Rows.Add(row)

            For j As Integer = 0 To tableColumnCount - 1
                Dim para As New Paragraph(document, String.Format("Cell {0}-{1}", i + 1, j + 1))

                row.Cells.Add(New TableCell(document, para))
            Next
        Next

        document.Sections.Add(New Section(document, table))

        document.Save("Simple Table.docx")

    End Sub

End Module