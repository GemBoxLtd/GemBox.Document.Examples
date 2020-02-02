Imports GemBox.Document

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As New DocumentModel()

        ' Built-in styles can be created using Style.CreateStyle method.
        Dim titleStyle = DirectCast(Style.CreateStyle(StyleTemplateType.Title, document), ParagraphStyle)

        ' We can also create our own custom styles.
        Dim emphasisStyle As New CharacterStyle("Emphasis")
        emphasisStyle.CharacterFormat.Italic = True

        ' To use styles, we first must add them to the document.
        document.Styles.Add(titleStyle)
        document.Styles.Add(emphasisStyle)

        ' Or we can use a utility method to get a built-in style or create and add a new one in a single statement.
        Dim strongStyle = DirectCast(document.Styles.GetOrAdd(StyleTemplateType.Strong), CharacterStyle)

        document.Sections.Add(
            New Section(document,
                New Paragraph(document, "Title (Title style)") With {.ParagraphFormat = New ParagraphFormat() With {.Style = titleStyle}},
                New Paragraph(document,
                    New Run(document, "Text is written using Emphasis style.") With {.CharacterFormat = New CharacterFormat() With {.Style = emphasisStyle}},
                    New SpecialCharacter(document, SpecialCharacterType.LineBreak),
                    New Run(document, "Text is written using Strong style.") With {.CharacterFormat = New CharacterFormat() With {.Style = strongStyle}})))

        document.Save("Styles.docx")

    End Sub
End Module