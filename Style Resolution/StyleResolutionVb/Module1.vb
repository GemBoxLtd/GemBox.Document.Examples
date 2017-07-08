Imports System
Imports GemBox.Document

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As DocumentModel = New DocumentModel

        ' Default font size is 12
        document.DefaultCharacterFormat.Size = 12

        Dim largeSize As New CharacterStyle("Large Font Size") With {
            .CharacterFormat = New CharacterFormat() With {
                .Size = 24
            }
        }

        ' Runs with the following style will have font size 24
        document.Styles.Add(largeSize)

        ' Fill section with test data.
        document.Sections.Add(
            New Section(document,
                New Paragraph(document,
                    New Run(document, "GemBox.Document component") With {
                        .CharacterFormat = New CharacterFormat() With {
                            .Style = largeSize
                        }
                    }),
                New Paragraph(document,
                    New Run(document, "GemBox.") With {
                        .CharacterFormat = New CharacterFormat() With {
                            .Style = largeSize,
                            .Size = 16
                        }
                    },
                    New Run(document, "Document") With {
                        .CharacterFormat = New CharacterFormat() With {
                            .Size = 16
                        }
                    },
                    New Run(document, " is a .NET component that enables developers to read, write, convert and print document files (DOCX, DOC, PDF, HTML, XPS, RTF and TXT) from .NET applications in a simple and efficient way.")
                )))

        ' Fill section with results.
        Dim section As New Section(document)
        Dim para As New Paragraph(document)
        section.Blocks.Add(para)

        For Each run As Run In document.GetChildElements(True, ElementType.Run)

            para.Inlines.Add(
                New Run(document, "Font size: " & run.CharacterFormat.Size & " points, Text: '" & run.Text & "'"))

            para.Inlines.Add(
                New SpecialCharacter(document, SpecialCharacterType.LineBreak))

        Next

        document.Sections.Add(section)

        document.Save("Style Resolution.docx")

    End Sub

End Module