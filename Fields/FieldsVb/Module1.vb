Imports System
Imports GemBox.Document

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As DocumentModel = New DocumentModel

        ' 1st field: { DATE }
        ' 2nd field: { DATE \@ "dddd, MMMM dd, yyyy"  \* MERGEFORMAT }
        ' 3rd field: { DATE  \@ "M/d/yyyy h:mm:ss am/pm"  \* MERGEFORMAT }
        document.Sections.Add(
            New Section(document,
                New Paragraph(document, "Press Alt + F9 to see field codes!"),
                New Paragraph(document,
                    New Run(document, "Date: "),
                    New Field(document, FieldType.[Date]),
                    New SpecialCharacter(document, SpecialCharacterType.LineBreak),
                    New Run(document, "Date (formatted): "),
                    New Field(document, FieldType.Date, "\@ ""dddd, MMMM dd, yyyy""  \* MERGEFORMAT"),
                    New SpecialCharacter(document, SpecialCharacterType.LineBreak),
                    New Run(document, "Date & Time (formatted): "),
                    New Field(document, FieldType.Date, " \@ ""M/d/yyyy h:mm:ss am/pm""  \* MERGEFORMAT"))))

        document.Save("Fields.docx")

    End Sub

End Module