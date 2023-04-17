Imports GemBox.Document
Imports GemBox.Document.Tables

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        ' If using the Professional version, remove this FreeLimitReached event handler.
        AddHandler ComponentInfo.FreeLimitReached, Sub(sender, e) e.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial

        Dim document As New DocumentModel()
        document.DefaultParagraphFormat.SpaceAfter = 0

        Dim section As New Section(document)
        document.Sections.Add(section)

        section.Blocks.Add(New Paragraph(document, "Built-in table style (List Table 2 - Accent 2):"))

        ' Create and add a built-in table style.
        Dim builtinTableStyle = DirectCast(Style.CreateStyle(StyleTemplateType.ListTable2Accent2, document), TableStyle)
        document.Styles.Add(builtinTableStyle)

        ' Modify built-in table style.
        builtinTableStyle.TableFormat.RowBandSize = 3

        ' Create a table and assign built-in table style to it.
        Dim tableWithBuiltinStyle As New Table(document, 10, 5,
            Function(r, c) New TableCell(document, New Paragraph(document, $"Cell ({r + 1},{c + 1})")))
        section.Blocks.Add(tableWithBuiltinStyle)
        tableWithBuiltinStyle.TableFormat.PreferredWidth = New TableWidth(100, TableWidthUnit.Percentage)
        tableWithBuiltinStyle.TableFormat.Style = builtinTableStyle
        tableWithBuiltinStyle.TableFormat.StyleOptions = TableStyleOptions.FirstRow Or TableStyleOptions.FirstColumn Or TableStyleOptions.BandedRows

        section.Blocks.Add(New Paragraph(document, "Custom table style (GemBox Table):") With {.ParagraphFormat = New ParagraphFormat() With {.SpaceBefore = 24}})

        ' Create and add a custom table style.
        Dim customTableStyle As New TableStyle("GemBox Table")
        document.Styles.Add(customTableStyle)

        ' Set table style format.
        customTableStyle.TableFormat.Borders.SetBorders(MultipleBorderTypes.All, BorderStyle.Single, New Color(79, 129, 189), 1)
        customTableStyle.ParagraphFormat.Alignment = HorizontalAlignment.Center

        ' Set table style conditional format for first row.
        Dim firstRowFormat = customTableStyle.ConditionalFormats(TableStyleFormatType.FirstRow)
        firstRowFormat.CharacterFormat.Bold = True
        firstRowFormat.CharacterFormat.FontColor = Color.White
        firstRowFormat.ParagraphFormat.SpaceAfter = 12
        firstRowFormat.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Single, New Color(79, 129, 189), 2.5)
        firstRowFormat.CellFormat.BackgroundColor = New Color(192, 80, 77)

        ' Set table style conditional format for last row.
        Dim lastRowFormat = customTableStyle.ConditionalFormats(TableStyleFormatType.LastRow)
        lastRowFormat.CharacterFormat.Bold = True
        lastRowFormat.ParagraphFormat.SpaceBefore = 12
        lastRowFormat.CellFormat.Borders.SetBorders(MultipleBorderTypes.Top, BorderStyle.Double, New Color(79, 129, 189), 1.25)
        lastRowFormat.CellFormat.BackgroundColor = New Color(211, 223, 238)

        ' Create a table and assign custom table style to it.
        Dim tableWithCustomStyle As New Table(document, 10, 5,
            Function(r, c) New TableCell(document, New Paragraph(document, $"Cell ({r + 1},{c + 1})")))
        section.Blocks.Add(tableWithCustomStyle)
        tableWithCustomStyle.TableFormat.PreferredWidth = New TableWidth(100, TableWidthUnit.Percentage)
        tableWithCustomStyle.TableFormat.Style = customTableStyle
        tableWithCustomStyle.TableFormat.StyleOptions = TableStyleOptions.FirstRow Or TableStyleOptions.LastRow

        document.Save("Table Styles.docx")

    End Sub
End Module