Imports System.Diagnostics
Imports System.IO
Imports System.Windows
Imports System.Windows.Documents
Imports System.Windows.Input
Imports GemBox.Document
Imports Microsoft.Win32

Namespace WpfRichTextEditor

    Partial Public Class MainWindow
        Inherits Window

        Public Sub New()

            InitializeComponent()

            ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        End Sub

        Private Sub Open(sender As Object, e As ExecutedRoutedEventArgs)

            Dim dialog = New OpenFileDialog() With {
                .AddExtension = True,
                .Filter =
                    "All Documents (*.docx;*.docm;*.doc;*.dotx;*.dotm;*.dot;*.htm;*.html;*.rtf;*.txt)|*.docx;*.docm;*.dotx;*.dotm;*.doc;*.dot;*.htm;*.html;*.rtf;*.txt|" +
                    "Word Documents (*.docx)|*.docx|" +
                    "Word Macro-Enabled Documents (*.docm)|*.docm|" +
                    "Word 97-2003 Documents (*.doc)|*.doc|" +
                    "Word Templates (*.dotx)|*.dotx|" +
                    "Word Macro-Enabled Templates (*.dotm)|*.dotm|" +
                    "Word 97-2003 Templates (*.dot)|*.dot|" +
                    "Web Pages (*.htm;*.html)|*.htm;*.html|" +
                    "Rich Text Format (*.rtf)|*.rtf|" +
                    "Text Files (*.txt)|*.txt"
            }

            If dialog.ShowDialog() = True Then
                Using stream = New MemoryStream()

                    ' Convert input file to RTF stream.
                    DocumentModel.Load(dialog.FileName).Save(stream, SaveOptions.RtfDefault)

                    stream.Position = 0

                    ' Load RTF stream into RichTextBox.
                    Dim textRange = New TextRange(Me.richTextBox.Document.ContentStart, Me.richTextBox.Document.ContentEnd)
                    textRange.Load(stream, DataFormats.Rtf)

                End Using
            End If

        End Sub

        Private Sub Save(sender As Object, e As ExecutedRoutedEventArgs)

            Dim dialog = New SaveFileDialog() With {
                .AddExtension = True,
                .Filter =
                    "Word Document (*.docx)|*.docx|" +
                    "Word Macro-Enabled Document (*.docm)|*.docm|" +
                    "Word Template (*.dotx)|*.dotx|" +
                    "Word Macro-Enabled Template (*.dotm)|*.dotm|" +
                    "PDF (*.pdf)|*.pdf|" +
                    "XPS Document (*.xps)|*.xps|" +
                    "Web Page (*.htm;*.html)|*.htm;*.html|" +
                    "Single File Web Page (*.mht;*.mhtml)|*.mht;*.mhtml|" +
                    "Rich Text Format (*.rtf)|*.rtf|" + "Plain Text (*.txt)|*.txt|" +
                    "Image (*.png;*.jpg;*.jpeg;*.gif;*.bmp;*.tif;*.tiff;*.wdp)|*.png;*.jpg;*.jpeg;*.gif;*.bmp;*.tif;*.tiff;*.wdp"
            }

            If dialog.ShowDialog(Me) = True Then
                Using stream = New MemoryStream()

                    ' Save RichTextBox content to RTF stream.
                    Dim textRange = New TextRange(Me.richTextBox.Document.ContentStart, Me.richTextBox.Document.ContentEnd)
                    textRange.Save(stream, DataFormats.Rtf)

                    stream.Position = 0

                    ' Convert RTF stream to output format.
                    DocumentModel.Load(stream, LoadOptions.RtfDefault).Save(dialog.FileName)
                    Process.Start(dialog.FileName)

                End Using
            End If

        End Sub

        Private Sub Cut(sender As Object, e As ExecutedRoutedEventArgs)

            Me.Copy(sender, e)

            ' Clear selection.
            Me.richTextBox.Selection.Text = String.Empty

        End Sub

        Private Sub Copy(sender As Object, e As ExecutedRoutedEventArgs)

            Using stream = New MemoryStream()

                ' Save RichTextBox selection to RTF stream.
                Me.richTextBox.Selection.Save(stream, DataFormats.Rtf)

                stream.Position = 0

                ' Save RTF stream to clipboard.
                DocumentModel.Load(stream, LoadOptions.RtfDefault).Content.SaveToClipboard()

            End Using

        End Sub

        Private Sub Paste(sender As Object, e As ExecutedRoutedEventArgs)

            Using stream = New MemoryStream()

                ' Save RichTextBox content to RTF stream.
                Dim textRange = New TextRange(Me.richTextBox.Document.ContentStart, Me.richTextBox.Document.ContentEnd)
                textRange.Save(stream, DataFormats.Rtf)

                stream.Position = 0

                ' Load document from RTF stream and prepend or append clipboard content to it.
                Dim document = DocumentModel.Load(stream, LoadOptions.RtfDefault)
                Dim position = If(DirectCast(e.Parameter, String) = "prepend", document.Content.Start, document.Content.End)
                position.LoadFromClipboard()

                stream.Position = 0

                ' Save document to RTF stream.
                document.Save(stream, SaveOptions.RtfDefault)

                stream.Position = 0

                ' Load RTF stream into RichTextBox.
                textRange.Load(stream, DataFormats.Rtf)

            End Using

        End Sub

        Private Sub CanSave(sender As Object, e As CanExecuteRoutedEventArgs)

            If Me.richTextBox IsNot Nothing Then
                Dim document = Me.richTextBox.Document
                Dim startPosition = document.ContentStart.GetNextInsertionPosition(LogicalDirection.Forward)
                Dim endPosition = document.ContentEnd.GetNextInsertionPosition(LogicalDirection.Backward)
                e.CanExecute = startPosition IsNot Nothing AndAlso endPosition IsNot Nothing AndAlso startPosition.CompareTo(endPosition) < 0
            Else
                e.CanExecute = False
            End If

        End Sub

        Private Sub CanCut(sender As Object, e As CanExecuteRoutedEventArgs)

            e.CanExecute = Me.richTextBox IsNot Nothing AndAlso Not Me.richTextBox.Selection.IsEmpty

        End Sub

        Private Sub CanCopy(sender As Object, e As CanExecuteRoutedEventArgs)

            e.CanExecute = Me.richTextBox IsNot Nothing AndAlso Not Me.richTextBox.Selection.IsEmpty

        End Sub

        Private Sub CanPaste(sender As Object, e As CanExecuteRoutedEventArgs)

            e.CanExecute = Me.richTextBox IsNot Nothing AndAlso Me.richTextBox.IsKeyboardFocused

        End Sub

    End Class

End Namespace