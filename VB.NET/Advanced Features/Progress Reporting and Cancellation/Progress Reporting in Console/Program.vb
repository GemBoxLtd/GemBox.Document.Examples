Imports GemBox.Document
Imports System

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Console.WriteLine("Creating document")

        ' Create large document.
        Dim document As New DocumentModel()
        Dim section As New Section(document)
        document.Sections.Add(section)
        For i As Integer = 0 To 10000
            section.Blocks.Add(New Paragraph(document, i.ToString()))
        Next

        ' Create save options.
        Dim saveOptions = New DocxSaveOptions()
        AddHandler saveOptions.ProgressChanged,
            Sub(eventSender, args)
                Console.WriteLine($"Progress changed - {args.ProgressPercentage}%")
            End Sub

        ' Save document.
        document.Save("document.docx", saveOptions)

    End Sub

End Module