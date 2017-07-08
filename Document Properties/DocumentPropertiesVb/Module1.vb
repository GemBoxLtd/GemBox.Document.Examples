Imports System
Imports GemBox.Document

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As DocumentModel = DocumentModel.Load("Reading.docx")

        Dim builtInPara As New Paragraph(document,
                        New Run(document, "Built-in document properties:"),
                        New SpecialCharacter(document, SpecialCharacterType.LineBreak))
        builtInPara.ParagraphFormat.Alignment = HorizontalAlignment.Left

        For Each docProp In document.DocumentProperties.BuiltIn

            builtInPara.Inlines.Add(
        New Run(document, String.Format("{0}: {1}", docProp.Key, docProp.Value)))

            builtInPara.Inlines.Add(
        New SpecialCharacter(document, SpecialCharacterType.LineBreak))

        Next

        Dim customPropPara As New Paragraph(document,
                        New Run(document, "Custom document properties:"),
                        New SpecialCharacter(document, SpecialCharacterType.LineBreak))
        customPropPara.ParagraphFormat.Alignment = HorizontalAlignment.Left

        For Each docProp In document.DocumentProperties.Custom

            customPropPara.Inlines.Add(
        New Run(document, String.Format("{0}: {1} (Type: {2})", docProp.Key, docProp.Value, docProp.Value.GetType())))

            customPropPara.Inlines.Add(
        New SpecialCharacter(document, SpecialCharacterType.LineBreak))

        Next

        document.Sections.Clear()
        document.Sections.Add(New Section(document, builtInPara, customPropPara))

        document.Save("Document Properties.docx")

    End Sub

End Module