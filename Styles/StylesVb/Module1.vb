Imports System
Imports GemBox.Document

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As DocumentModel = New DocumentModel

        ' Built-in styles can be created using Style.CreateStyle() method.
        Dim title = DirectCast(Style.CreateStyle(StyleTemplateType.Title, document), ParagraphStyle)

        ' We can also create our own (custom) styles.
        Dim emphasis As New CharacterStyle("Emphasis")
        emphasis.CharacterFormat.Italic = True

        ' First add style to the document, then use it.
        document.Styles.Add(title)
        document.Styles.Add(emphasis)

        document.Sections.Add(
            New Section(document,
                New Paragraph(document, "Title (Title style)") With {
                    .ParagraphFormat = New ParagraphFormat() With {
                        .Style = title
                    }
                },
                New Paragraph(document,
                    New Run(document, "Text is written using Strong style.") With {
                        .CharacterFormat = New CharacterFormat() With {
                            .Style = document.Styles.GetOrAdd(StyleTemplateType.Strong)
                        }
                    },
                    New SpecialCharacter(document, SpecialCharacterType.LineBreak),
                    New Run(document, "Text is written using Emphasis style.") With {
                        .CharacterFormat = New CharacterFormat() With {
                            .Style = emphasis
                        }
                    })))

        document.Save("Styles.docx")

    End Sub

End Module