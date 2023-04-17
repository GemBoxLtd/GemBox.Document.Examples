Imports System
Imports GemBox.Document

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        ' If using the Professional version, remove this FreeLimitReached event handler.
        AddHandler ComponentInfo.FreeLimitReached, Sub(sender, e) e.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial

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