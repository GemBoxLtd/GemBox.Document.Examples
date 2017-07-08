Imports System
Imports System.Linq
Imports System.Text
Imports GemBox.Document

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As DocumentModel = DocumentModel.Load("Reading.docx")

        Dim numberOfSections As Integer = document.Sections.Count
        Dim numberOfParagraphs As Integer = document.GetChildElements(True, ElementType.Paragraph).Count()
        Dim numberOfRunsAndFields As Integer = document.GetChildElements(True, ElementType.Run, ElementType.Field).Count()
        Dim numberOfInlines As Integer = document.GetChildElements(True).OfType(Of Inline)().Count()

        Dim elements As Integer = document.Sections(0).GetChildElements(True).Count()

        Dim sb As New StringBuilder()
        sb.AppendLine("File has:")
        sb.AppendLine(numberOfSections & " section")
        sb.AppendLine(numberOfParagraphs & " paragraphs")
        sb.AppendLine(numberOfRunsAndFields & " runs and fields")
        sb.AppendLine(numberOfInlines & " inlines")
        sb.AppendLine("First section contains " & elements & " elements")

        Console.WriteLine(sb.ToString())

    End Sub

End Module