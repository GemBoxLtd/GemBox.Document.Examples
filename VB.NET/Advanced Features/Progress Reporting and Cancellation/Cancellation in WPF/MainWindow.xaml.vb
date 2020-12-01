Imports System.Threading
Imports GemBox.Document

Class MainWindow

    Private Property cancellationRequested As Boolean

    Public Sub New()
        ' If using Professional version, put your serial key below
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")
        ' Use Trial Mode
        AddHandler ComponentInfo.FreeLimitReached,
            Sub(eventSender, args)
                args.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial
            End Sub

        InitializeComponent()
    End Sub

    Private Async Sub loadButton_Click(sender As Object, e As RoutedEventArgs) Handles loadButton.Click
        ' Capture the current context on the UI thread
        Dim context = SynchronizationContext.Current

        ' Create load options
        Dim loadOptions = New DocxLoadOptions()
        AddHandler loadOptions.ProgressChanged,
            Sub(eventSender, args)
                ' Show progress
                context.Post(
                    Sub(progressPercentage)
                        Me.progressBar.Value = CType(progressPercentage, Integer)
                    End Sub, args.ProgressPercentage)

                ' Cancel if requested
                If Me.cancellationRequested Then
                    args.CancelOperation()
                End If
            End Sub

        Try
            Dim file = Await Threading.Tasks.Task.Run(
                Function() As DocumentModel
                    Return DocumentModel.Load("LargeDocument.docx", loadOptions)
                End Function)
        Catch ex As OperationCanceledException
            ' Operation cancelled
        End Try
    End Sub

    Private Sub cancelButton_Click(sender As Object, e As RoutedEventArgs) Handles cancelButton.Click
        Me.cancellationRequested = True
    End Sub
End Class
