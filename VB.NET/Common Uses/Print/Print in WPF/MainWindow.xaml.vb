Imports GemBox.Document
Imports Microsoft.Win32
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Xps.Packaging

Partial Public Class MainWindow
    Inherits Window

    Dim document As DocumentModel

    Public Sub New()
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")
        InitializeComponent()
    End Sub

    Private Sub LoadFileBtn_Click(sender As Object, e As RoutedEventArgs)

        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Filter =
            "DOCX files (*.docx, *.dotx, *.docm, *.dotm)|*.docx;*.dotx;*.docm;*.dotm" &
            "|DOC files (*.doc, *.dot)|*.doc;*.dot" &
            "|RTF files (*.rtf)|*.rtf" &
            "|HTML files (*.html, *.htm)|*.html;*.htm" &
            "|PDF files (*.pdf)|*.pdf" &
            "|Word XML files (*.xml)|*.xml" &
            "|TXT files (*.txt)|*.txt"

        If (openFileDialog.ShowDialog() = True) Then
            Me.document = DocumentModel.Load(openFileDialog.FileName)
            Me.ShowPrintPreview()
        End If

    End Sub

    Private Sub PrintFileBtn_Click(sender As Object, e As RoutedEventArgs)

        If document Is Nothing Then Return

        Dim printDialog As New PrintDialog() With {.UserPageRangeEnabled = True}
        If (printDialog.ShowDialog() = True) Then

            Dim printOptions As New PrintOptions(printDialog.PrintTicket.GetXmlStream())

            printOptions.FromPage = printDialog.PageRange.PageFrom - 1
            printOptions.ToPage = If(printDialog.PageRange.PageTo = 0, Integer.MaxValue, printDialog.PageRange.PageTo - 1)

            Me.document.Print(printDialog.PrintQueue.FullName, printOptions)
        End If

    End Sub

    Private Sub ShowPrintPreview()

        Dim xpsDocument As XpsDocument = document.ConvertToXpsDocument(SaveOptions.XpsDefault)

        ' Note, XpsDocument must stay referenced so that DocumentViewer can access additional resources from it.
        ' Otherwise, GC will collect/dispose XpsDocument and DocumentViewer will no longer work.
        Me.DocViewer.Tag = xpsDocument
        Me.DocViewer.Document = xpsDocument.GetFixedDocumentSequence()

    End Sub

End Class