Imports System
Imports System.Text
Imports GemBox.Document

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As DocumentModel = DocumentModel.Load("Reading.docx")

        Dim sb As New StringBuilder()

        For Each paragraph As Paragraph In document.GetChildElements(True, ElementType.Paragraph)
            For Each run As Run In paragraph.GetChildElements(True, ElementType.Run)
                Dim isBold As Boolean = run.CharacterFormat.Bold
                Dim text As String = run.Text

                sb.AppendFormat("{0}{1}{2}", If(isBold, "<b>", ""), text, If(isBold, "</b>", ""))
            Next
            sb.AppendLine()
        Next

        Console.WriteLine(sb.ToString())

    End Sub

End Module