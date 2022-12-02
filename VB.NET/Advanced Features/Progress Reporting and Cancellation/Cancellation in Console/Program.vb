Imports System
Imports System.Diagnostics
Imports GemBox.Document

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")
        ' Use Trial Mode
        AddHandler ComponentInfo.FreeLimitReached,
            Sub(eventSender, args)
                args.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial
            End Sub

        ' Create document
        Dim document As New DocumentModel()
        Dim section As New Section(document)
        document.Sections.Add(section)
        For i As Integer = 0 To 10000
            section.Blocks.Add(New Paragraph(document, i.ToString()))
        Next

        Dim stopwatch = New Stopwatch()
        stopwatch.Start()

        ' Create save options
        Dim saveOptions = New DocxSaveOptions()
        AddHandler saveOptions.ProgressChanged,
            Sub(eventSender, args)
                ' Cancel operation after five seconds
                If stopwatch.Elapsed.Seconds >= 5 Then
                    args.CancelOperation()
                End If
            End Sub

        Try
            document.Save("Cancellation.docx", saveOptions)
            Console.WriteLine("Operation fully finished")
        Catch ex As OperationCanceledException
            Console.WriteLine("Operation was cancelled")
        End Try
    End Sub
End Module