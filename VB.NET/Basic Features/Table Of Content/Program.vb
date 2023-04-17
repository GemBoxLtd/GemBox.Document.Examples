Imports System.Linq
Imports GemBox.Document

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        ' If using the Professional version, remove this FreeLimitReached event handler.
        AddHandler ComponentInfo.FreeLimitReached, Sub(sender, e) e.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial

        Dim heading1Count As Integer = 3
        Dim heading2Count As Integer = 5
            
        Dim document As New DocumentModel()

        Dim section As New Section(document)
        document.Sections.Add(section)

        ' Create and add "Heading 1" style.
        Dim heading1Style = DirectCast(Style.CreateStyle(StyleTemplateType.Heading1, document), ParagraphStyle)
        document.Styles.Add(heading1Style)

        ' Create and add "Heading 2" style.
        Dim heading2Style = DirectCast(Style.CreateStyle(StyleTemplateType.Heading2, document), ParagraphStyle)
        document.Styles.Add(heading2Style)

        ' Create and add new TOC element.
        section.Blocks.Add(New TableOfEntries(document, FieldType.TOC))

        section.Blocks.Add(
            New Paragraph(document,
                New SpecialCharacter(document, SpecialCharacterType.PageBreak)))

        ' Add document content.
        For i As Integer = 0 To heading1Count - 1

            ' Add "Heading 1" paragraph with Level1 for OutlineLevel. 
            section.Blocks.Add(
                New Paragraph(document, $"Heading 1 ({i + 1})") With {.ParagraphFormat = New ParagraphFormat() With {.Style = heading1Style}})

            For j As Integer = 0 To heading2Count - 1

                ' Add "Heading 2" paragraph with Level2 for OutlineLevel.
                section.Blocks.Add(
                    New Paragraph(document, $"Heading 2 ({i + 1}-{j + 1})") With {.ParagraphFormat = New ParagraphFormat() With {.Style = heading2Style}})

                section.Blocks.Add(
                    New Paragraph(document, "This is a paragraph.\nIt has a default BodyText for OutlineLevel.\nIt won't be listed in TOC entries."))
            Next
        Next

        ' Get and update TOC element.
        Dim toc = DirectCast(document.GetChildElements(True, ElementType.TableOfEntries).First(), TableOfEntries)
        toc.Update()

        ' Update TOC entries page numbers.
        ' This is not needed when saving to PDF, XPS and image format or when printing.
        ' Page numbers are automatically updated in that case.
        ' document.GetPaginator(New PaginatorOptions() With {.UpdateFields = True})

        document.Save("TOC.docx")

    End Sub
End Module