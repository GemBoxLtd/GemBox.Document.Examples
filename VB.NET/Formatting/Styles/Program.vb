Imports GemBox.Document
Imports System.Linq

Module Program

    Sub Main()
        Example1()
        Example2()
    End Sub

    Sub Example1()
        ' If using the Professional version, put your serial key below.
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

    Sub Example2()
        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As New DocumentModel()

        ' Document's default font size is 8pt.
        document.DefaultCharacterFormat.Size = 8

        ' Style's font size is 24pt.
        Dim largeFont As New CharacterStyle("Large Font") With {.CharacterFormat = New CharacterFormat() With {.Size = 24}}
        document.Styles.Add(largeFont)

        Dim section As New Section(document)
        document.Sections.Add(section)

        Dim paragraph As New Paragraph(document,
            New Run(document, "Large text that has 'Large Font' style.") With
            {
                .CharacterFormat = New CharacterFormat() With {.Style = largeFont}
            },
            New SpecialCharacter(document, SpecialCharacterType.LineBreak),
            New Run(document, "Medium text that has both style and direct formatting; direct formatting has precedence over style's formatting.") With
            {
                .CharacterFormat = New CharacterFormat() With {.Style = largeFont, .Size = 12}
            },
            New SpecialCharacter(document, SpecialCharacterType.LineBreak),
            New Run(document, "Small text that uses document's default formatting."))

        section.Blocks.Add(paragraph)

        ' Write elements resolved font size values.
        For Each run As Run In document.GetChildElements(True, ElementType.Run).ToArray()
            section.Blocks.Add(New Paragraph(document, $"Font size: {run.CharacterFormat.Size} points. Text: {run.Text}"))
        Next

        document.Save("Style Resolution.docx")
    End Sub

End Module
