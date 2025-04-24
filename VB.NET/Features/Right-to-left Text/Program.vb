Imports GemBox.Document
Imports GemBox.Document.Tables

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document = DocumentModel.Load("RightToLeft.docx")

        ' Show line numbers on the right side of the page
        Dim pageSetup = document.Sections(0).PageSetup
        pageSetup.RightToLeft = true
        pageSetup.LineNumberRestartSetting = LineNumberRestartSetting.Continuous

        ' Create a new right-to-left paragraph
        Dim paragraph = new Paragraph(document)
        paragraph.ParagraphFormat.RightToLeft = true
        paragraph.Inlines.Add(new Run(document, "أخذ عن موالية الإمتعاض"))
        document.Sections(0).Blocks.Add(paragraph)

        ' Create a right-to-left table
        Dim table = new Table(document)
        table.TableFormat.RightToLeft = true
        table.TableFormat.PreferredWidth = new TableWidth(100, TableWidthUnit.Percentage)
        Dim row = new TableRow(document)
        table.Rows.Add(row)

        Dim firstCellPara = new Paragraph(document, "של תיבת תרומה מלא")
        firstCellPara.ParagraphFormat.RightToLeft = true
        row.Cells.Add(new TableCell(document, firstCellPara))

        Dim secondCellPara = new Paragraph(document, "200")
        row.Cells.Add(new TableCell(document, secondCellPara))

        document.Sections(0).Blocks.Add(table)

        document.Save("RightToLeft.pdf")
    End Sub

End Module
