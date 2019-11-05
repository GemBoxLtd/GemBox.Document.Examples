Imports System
Imports GemBox.Document

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As DocumentModel = DocumentModel.Load("Invoice.docx")

        ' Get content from each paragraph.
        For Each paragraph As Paragraph In document.GetChildElements(True, ElementType.Paragraph)
            Console.WriteLine($"Paragraph: {paragraph.Content.ToString()}")
        Next

        ' Get content from each bold run.
        For Each run As Run In document.GetChildElements(True, ElementType.Run)
            If run.CharacterFormat.Bold Then
                Console.WriteLine($"Bold run: {run.Content.ToString()}")
            End If
        Next

    End Sub
End Module