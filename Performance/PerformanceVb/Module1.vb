Imports System
Imports System.Diagnostics
Imports GemBox.Document

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        ' If sample exceeds Free version limitations then continue as trial version: 
        ' https://www.gemboxsoftware.com/Document/help/html/Evaluation_and_Licensing.htm
        AddHandler ComponentInfo.FreeLimitReached, Sub(sender, e) e.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial

        Console.WriteLine("Performance sample:")
        Console.WriteLine()

        Dim stopwatch As New Stopwatch()
        stopwatch.Start()

        Dim document As DocumentModel = DocumentModel.Load("Template.docx", LoadOptions.DocxDefault)

        Console.WriteLine("Load file (seconds): " & stopwatch.Elapsed.TotalSeconds)

        stopwatch.Reset()
        stopwatch.Start()

        Dim numberOfParagraphs As Integer = 0
        For Each item As Paragraph In document.GetChildElements(True, ElementType.Paragraph)
            numberOfParagraphs += 1
        Next

        Console.WriteLine("Iterate through " & numberOfParagraphs & " paragraphs (seconds): " & stopwatch.Elapsed.TotalSeconds)

        stopwatch.Reset()
        stopwatch.Start()

        document.Save("Report.docx")

        Console.WriteLine("Save file (seconds): " & stopwatch.Elapsed.TotalSeconds)

    End Sub

End Module
