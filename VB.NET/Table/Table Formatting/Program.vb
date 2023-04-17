Imports GemBox.Document
Imports GemBox.Document.Tables

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        ' If using the Professional version, remove this FreeLimitReached event handler.
        AddHandler ComponentInfo.FreeLimitReached, Sub(sender, e) e.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial

        Example1()
        Example2()

    End Sub

    Sub Example1()
        Dim document As New DocumentModel()

        Dim table As New Table(document)
        table.TableFormat.AutomaticallyResizeToFitContents = False

        ' By default Table has assigned "Table Grid" style, the same as when creating it in Microsoft Word.
        ' This base style defines borders which can be removed with the following.
        table.TableFormat.Style.TableFormat.Borders.ClearBorders()

        ' Add columns with specified width.
        table.Columns.Add(New TableColumn(60))
        table.Columns.Add(New TableColumn(120))
        table.Columns.Add(New TableColumn(180))

        ' Add rows with specified height.
        table.Rows.Add(New TableRow(document) With {.RowFormat = New TableRowFormat() With {.Height = New TableRowHeight(30, TableRowHeightRule.AtLeast)}})
        table.Rows.Add(New TableRow(document) With {.RowFormat = New TableRowFormat() With {.Height = New TableRowHeight(60, TableRowHeightRule.AtLeast)}})
        table.Rows.Add(New TableRow(document) With {.RowFormat = New TableRowFormat() With {.Height = New TableRowHeight(90, TableRowHeightRule.AtLeast)}})

        For r = 0 To 2
            For c = 0 To 2

                ' Add cell.
                Dim cell As New TableCell(document)
                table.Rows(r).Cells.Add(cell)

                ' Set cell's vertical alignment.
                cell.CellFormat.VerticalAlignment = CType(r, VerticalAlignment)

                ' Add cell content.
                Dim paragraph As New Paragraph(document, $"Cell ({r + 1},{c + 1})")
                cell.Blocks.Add(paragraph)

                ' Set cell content's horizontal alignment.
                paragraph.ParagraphFormat.Alignment = CType(c, HorizontalAlignment)

                If (r + c) Mod 2 = 0 Then

                    ' Set cell's background and borders.
                    cell.CellFormat.BackgroundColor = New Color(255, 242, 204)
                    cell.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Double, Color.Red, 1)

                End If
            Next
        Next

        document.Sections.Add(New Section(document, table))

        document.Save("Table Formatting.docx")
    End Sub

    Sub Example2()
        Dim document As New DocumentModel()

        Dim section As New Section(document)
        document.Sections.Add(section)

        section.Blocks.Add(New Paragraph(document,
            New Run(document, "Keep table on same page.") With {.CharacterFormat = New CharacterFormat With {.Size = 36}},
            New SpecialCharacter(document, SpecialCharacterType.LineBreak),
            New Run(document, "This paragraph has a large spacing before to occupy most of the page.") With {.CharacterFormat = New CharacterFormat With {.Size = 14}}) With
        {.ParagraphFormat = New ParagraphFormat With {.SpaceBefore = 400}})

        Dim table As New Table(document, 20, 4)
        table.TableFormat.PreferredWidth = New TableWidth(100, TableWidthUnit.Percentage)
        section.Blocks.Add(table)

        ' If you were to save a document at this point, you'd notice that the last few rows don't fit on the same page.
        ' In other words, the table rows break across the first and second page.
        'document.Save("TableOnTwoPages.docx")

        ' To prevent the table breaking across two pages, you need to set the KeepWithNext formatting.
        For Each cell As TableCell In table.GetChildElements(True, ElementType.TableCell)
            ' Cell should have at least one paragraph.
            If cell.Blocks.Count = 0 Then cell.Blocks.Add(New Paragraph(cell.Document))

            For Each paragraph As Paragraph In cell.GetChildElements(True, ElementType.Paragraph)
                paragraph.ParagraphFormat.KeepWithNext = True
            Next
        Next

        document.Save("TableOnOnePage.docx")
    End Sub

End Module