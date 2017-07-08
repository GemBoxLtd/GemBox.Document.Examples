Imports System
Imports System.Text
Imports GemBox.Document

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As DocumentModel = DocumentModel.Load("Reading.docx")

        Dim sb = New StringBuilder()

        ' Get content from each paragraph
        For Each paragraph As Paragraph In document.GetChildElements(True, ElementType.Paragraph)
            sb.AppendFormat("Paragraph: {0}", paragraph.Content.ToString())
            sb.AppendLine()
        Next

        ' Get content from each bold run
        For Each run As Run In document.GetChildElements(True, ElementType.Run)
            If (run.CharacterFormat.Bold) Then
                sb.AppendFormat("Bold run: {0}", run.Content.ToString())
                sb.AppendLine()
            End If
        Next

        Console.WriteLine(sb.ToString())

    End Sub

End Module