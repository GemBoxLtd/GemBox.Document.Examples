Imports System
Imports GemBox.Document

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As DocumentModel = New DocumentModel

        Dim section As New Section(document,
            New Paragraph(document,
                New Run(document, "First page"),
                New SpecialCharacter(document, SpecialCharacterType.PageBreak)),
            New Paragraph(document,
                New Run(document, "Even page"),
                New SpecialCharacter(document, SpecialCharacterType.PageBreak)),
            New Paragraph(document,
                New Run(document, "Odd page"),
                New SpecialCharacter(document, SpecialCharacterType.PageBreak)),
            New Paragraph(document, "Even page"))

        section.HeadersFooters.Add(
            New HeaderFooter(document, HeaderFooterType.HeaderFirst,
                New Paragraph(document,
                    New Run(document, "First Header"))))

        ' Add page number.
        section.HeadersFooters.Add(
            New HeaderFooter(document, HeaderFooterType.FooterFirst,
                New Paragraph(document,
                    New Run(document, "First Footer")),
                New Paragraph(document,
                    New Field(document, FieldType.Page)) With {
                .ParagraphFormat = New ParagraphFormat() With {
                    .Alignment = HorizontalAlignment.Right
         }}))

        section.HeadersFooters.Add(
            New HeaderFooter(document, HeaderFooterType.HeaderDefault,
                New Paragraph(document,
                    New Run(document, "Default Header"))))

        section.HeadersFooters.Add(
            New HeaderFooter(document, HeaderFooterType.FooterDefault,
                New Paragraph(document,
                    New Run(document, "Default Footer")),
                New Paragraph(document,
                    New Field(document, FieldType.Page)) With {
                .ParagraphFormat = New ParagraphFormat() With {
                    .Alignment = HorizontalAlignment.Right
         }}))

        section.HeadersFooters.Add(
            New HeaderFooter(document, HeaderFooterType.HeaderEven,
                New Paragraph(document,
                    New Run(document, "Even Header"))))

        section.HeadersFooters.Add(
            New HeaderFooter(document, HeaderFooterType.FooterEven,
                New Paragraph(document,
                    New Run(document, "Even Footer")),
                New Paragraph(document,
                    New Field(document, FieldType.Page)) With {
                .ParagraphFormat = New ParagraphFormat() With {
                    .Alignment = HorizontalAlignment.Right
         }}))

        document.Sections.Add(section)

        document.Save("Header and Footer.docx")

    End Sub

End Module