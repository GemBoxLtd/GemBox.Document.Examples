Imports GemBox.Document

Module Program

    Sub Main()
        Example1()
        Example2()
    End Sub

    Sub Example1()
        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As New DocumentModel()
        document.DefaultCharacterFormat.Size = 48

        Dim section As New Section(document,
            New Paragraph(document, "First page"),
            New Paragraph(document, New SpecialCharacter(document, SpecialCharacterType.PageBreak)),
            New Paragraph(document, "Even page"),
            New Paragraph(document, New SpecialCharacter(document, SpecialCharacterType.PageBreak)),
            New Paragraph(document, "Odd page"),
            New Paragraph(document, New SpecialCharacter(document, SpecialCharacterType.PageBreak)),
            New Paragraph(document, "Even page"))

        document.Sections.Add(section)

        ' Add default (odd) header.
        section.HeadersFooters.Add(
            New HeaderFooter(document, HeaderFooterType.HeaderDefault,
                New Paragraph(document, "Default Header")))

        ' Add default (odd) footer with page number.
        section.HeadersFooters.Add(
            New HeaderFooter(document, HeaderFooterType.FooterDefault,
                New Paragraph(document, "Default Footer")))

        ' Add first header.
        section.HeadersFooters.Add(
            New HeaderFooter(document, HeaderFooterType.HeaderFirst,
                New Paragraph(document, "First Header")))

        ' Add first footer.
        section.HeadersFooters.Add(
            New HeaderFooter(document, HeaderFooterType.FooterFirst,
                New Paragraph(document, "First Footer"),
                New Paragraph(document,
                    New Field(document, FieldType.Page),
                    New Run(document, " of "),
                    New Field(document, FieldType.NumPages)) With
                {
                    .ParagraphFormat = New ParagraphFormat() With {.Alignment = HorizontalAlignment.Right}
                }))

        ' Add even header.
        section.HeadersFooters.Add(
            New HeaderFooter(document, HeaderFooterType.HeaderEven,
                New Paragraph(document, "Even Header")))

        ' Add even footer with page number.
        section.HeadersFooters.Add(
            New HeaderFooter(document, HeaderFooterType.FooterEven,
                New Paragraph(document, "Even Footer"),
                New Paragraph(document,
                    New Field(document, FieldType.Page),
                    New Run(document, " of "),
                    New Field(document, FieldType.NumPages)) With
                {
                    .ParagraphFormat = New ParagraphFormat() With {.Alignment = HorizontalAlignment.Right}
                }))

        document.Save("Header and Footer.docx")
    End Sub

    Sub Example2()
        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As New DocumentModel()
        document.DefaultCharacterFormat.Size = 28

        ' Create sections.
        Dim section1 As New Section(document, New Paragraph(document, "First section content") With
        {.ParagraphFormat = New ParagraphFormat() With {.SpaceBefore = 40, .SpaceAfter = 0}})
        Dim section2 As New Section(document, New Paragraph(document, "Second section content") With
        {.ParagraphFormat = New ParagraphFormat() With {.SpaceBefore = 40, .SpaceAfter = 0}})
        Dim section3 As New Section(document, New Paragraph(document, "Third section content") With
        {.ParagraphFormat = New ParagraphFormat() With {.SpaceBefore = 40, .SpaceAfter = 0}})

        ' Add sections to the document.
        document.Sections.Add(section1)
        document.Sections.Add(section2)
        document.Sections.Add(section3)

        ' Create a header in the first section.
        section1.HeadersFooters.Add(
            New HeaderFooter(document, HeaderFooterType.HeaderDefault,
                New Paragraph(document, "Shared Header (linked across sections)")))

        ' Link headers in the second and third sections to the first section header.
        section2.HeadersFooters.SetLinkedToPrevious(HeaderFooterType.HeaderDefault, True)
        section3.HeadersFooters.SetLinkedToPrevious(HeaderFooterType.HeaderDefault, True)

        document.Save("Linked To Previous.docx")
    End Sub

End Module
