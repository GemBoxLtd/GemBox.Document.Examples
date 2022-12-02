Imports System.Windows
Imports System.Windows.Xps.Packaging ' Add reference for 'ReachFramework'.
Imports GemBox.Document

Namespace ExportToXpsDocument
    Partial Public Class MainWindow
        Inherits Window

        Private xpsDocument As XpsDocument

        Public Sub New()

            Me.InitializeComponent()

            ' If using the Professional version, put your serial key below.
            ComponentInfo.SetLicense("FREE-LIMITED-KEY")

            Me.SetDocumentViewer("Reading.docx")

        End Sub

        Private Sub SetDocumentViewer(path As String)

            Dim document = DocumentModel.Load(path)

            ' XpsDocument object must stay referenced so that DocumentViewer can access its resources.
            ' Otherwise, GC will collect/dispose XpsDocument and DocumentViewer will no longer work.
            Me.xpsDocument = document.ConvertToXpsDocument(SaveOptions.XpsDefault)

            Me.DocumentViewer.Document = Me.xpsDocument.GetFixedDocumentSequence()

        End Sub

    End Class
End Namespace