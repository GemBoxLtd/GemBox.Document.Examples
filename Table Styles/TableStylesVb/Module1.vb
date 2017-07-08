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

        ' Built-in table styles can be created using Style.CreateStyle() method.
        Dim lightShadingStyle = DirectCast(Style.CreateStyle(StyleTemplateType.LightShadingAccent2, document), TableStyle)

        ' We can also create our own (custom) table styles.
        Dim customTableStyle As New TableStyle("GemBox Table Style") With {
                .ParagraphFormat = New ParagraphFormat() With {
                    .SpaceAfter = 0,
                    .LineSpacing = 1,
                    .LineSpacingRule = LineSpacingRule.Multiple
                },
                .TableFormat = New TableFormat() With {
                    .RowBandSize = 2,
                    .ColumnBandSize = 1,
                    .IndentFromLeft = 0D,
                    .DefaultCellPadding = New Padding(108 / 20D, 0D)
                }
            }

        ' Set custom style table borders.
        customTableStyle.TableFormat.Borders.SetBorders(MultipleBorderTypes.All, BorderStyle.Single, New Color(79, 129, 189), 1)

        ' Table style conditional format - First row.
        Dim firstRowFormat = customTableStyle.ConditionalFormats(TableStyleFormatType.FirstRow)
        firstRowFormat.CharacterFormat.Bold = True
        firstRowFormat.CharacterFormat.FontColor = Color.White
        firstRowFormat.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Single, New Color(79, 129, 189), 1)
        firstRowFormat.CellFormat.Borders.SetBorders(MultipleBorderTypes.Bottom, BorderStyle.Single, New Color(79, 129, 189), 2.5)
        firstRowFormat.CellFormat.BackgroundColor = New Color(192, 80, 77)

        ' Table style conditional format - Last row.
        Dim lastRowFormat = customTableStyle.ConditionalFormats(TableStyleFormatType.LastRow)
        lastRowFormat.CharacterFormat.Bold = True
        lastRowFormat.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Single, New Color(79, 129, 189), 1)
        lastRowFormat.CellFormat.Borders.SetBorders(MultipleBorderTypes.Top, BorderStyle.Double, New Color(79, 129, 189), 1.25)

        ' Table style conditional format - Banded rows.
        Dim bandedRowFormat = customTableStyle.ConditionalFormats(TableStyleFormatType.OddBandedRows)
        bandedRowFormat.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Single, New Color(79, 129, 189), 1)
        bandedRowFormat.CellFormat.BackgroundColor = New Color(211, 223, 238)

        ' First add style to the document, then use it.
        document.Styles.Add(lightShadingStyle)
        document.Styles.Add(customTableStyle)

        document.Sections.Add(
            New Section(document,
                New Paragraph(document, "Built-in Table Style:"),
                New Table(document, tableRowCount, tableColumnCount,
                    Function(rowIndex As Integer, columnIndex As Integer) New TableCell(document, New Paragraph(document, New Run(document, String.Format("{0}-{1}", rowIndex, columnIndex))))) With {
                    .TableFormat = New TableFormat() With {
                        .PreferredWidth = New TableWidth(100, TableWidthUnit.Percentage),
                        .Style = lightShadingStyle,
                        .StyleOptions = TableStyleOptions.FirstRow Or TableStyleOptions.FirstColumn Or TableStyleOptions.BandedRows
                    }
                },
                New Paragraph(document),
                New Paragraph(document, "Custom Table Style:"),
                New Table(document, tableRowCount, tableColumnCount,
                    Function(rowIndex As Integer, columnIndex As Integer) New TableCell(document, New Paragraph(document, New Run(document, String.Format("{0}-{1}", rowIndex, columnIndex))))) With {
                    .TableFormat = New TableFormat() With {
                        .PreferredWidth = New TableWidth(100, TableWidthUnit.Percentage),
                        .Style = customTableStyle,
                        .StyleOptions = TableStyleOptions.FirstRow Or TableStyleOptions.LastRow Or TableStyleOptions.BandedRows
                    }
                }))

        document.Save("Table Styles.docx")

    End Sub

End Module