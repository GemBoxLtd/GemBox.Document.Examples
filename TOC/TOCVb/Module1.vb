Imports System
Imports System.Linq
Imports GemBox.Document

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As DocumentModel = New DocumentModel

        Dim heading1Count As Integer = 3
        Dim heading2Count As Integer = 5

        ' Create and add Heading 1 style.
        Dim heading1Style As ParagraphStyle = DirectCast(Style.CreateStyle(StyleTemplateType.Heading1, document), ParagraphStyle)
        document.Styles.Add(heading1Style)

        ' Create and add Heading 2 style.
        Dim heading2Style As ParagraphStyle = DirectCast(Style.CreateStyle(StyleTemplateType.Heading2, document), ParagraphStyle)
        document.Styles.Add(heading2Style)

        ' Create and add TOC style.
        Dim tocHeading As ParagraphStyle = DirectCast(Style.CreateStyle(StyleTemplateType.Heading1, document), ParagraphStyle)
        tocHeading.Name = "toc"
        tocHeading.ParagraphFormat.OutlineLevel = OutlineLevel.BodyText
        document.Styles.Add(tocHeading)

        Dim section As New Section(document)
        document.Sections.Add(section)

        ' Add TOC title.
        section.Blocks.Add(
            New Paragraph(document, "Contents") With {
                .ParagraphFormat = New ParagraphFormat() With {.Style = tocHeading}})

        ' Create and add new TOC.
        section.Blocks.Add(
            New TableOfEntries(document, FieldType.TOC))

        section.Blocks.Add(
            New Paragraph(document,
                New SpecialCharacter(document, SpecialCharacterType.PageBreak)))

        ' Add document content.
        For i As Integer = 0 To heading1Count - 1

            ' Heading 1
            section.Blocks.Add(
                New Paragraph(document, "Heading 1 (" & (i + 1) & ")") With {
                 .ParagraphFormat = New ParagraphFormat() With {
                     .Style = heading1Style}
                })

            For j As Integer = 0 To heading2Count - 1
                ' Heading 2
                section.Blocks.Add(
                    New Paragraph(document, String.Format("Heading 2 ({0}-{1})", i + 1, j + 1)) With {
                     .ParagraphFormat = New ParagraphFormat() With {
                         .Style = heading2Style}
                })

                ' Heading 2 content.
                section.Blocks.Add(
                    New Paragraph(document,
                        "GemBox.Document is a .NET component that enables developers to read, write, convert and print document files (DOCX, DOC, PDF, HTML, XPS, RTF and TXT) from .NET applications in a simple and efficient way."))
            Next
        Next

        ' Update TOC (TOC can be updated only after all document content is added).
        Dim toc = DirectCast(document.GetChildElements(True, ElementType.TableOfEntries).First(), TableOfEntries)
        toc.Update()

        ' Update TOC's page numbers.
        ' NOTE: This is not necessary when printing and saving to PDF, XPS or an image format.
        ' Page numbers are automatically updated in that case.
        document.GetPaginator(New PaginatorOptions() With {.UpdateFields = True})

        document.Save("TOC.docx")

    End Sub

End Module