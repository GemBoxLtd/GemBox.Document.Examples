Imports System.Windows.Xps.Packaging
Imports GemBox.Document

Class MainWindow

    Dim xpsDocument As XpsDocument

    Public Sub New()
        InitializeComponent()

        SetDocumentViewer(Me.DocumentViewer)
    End Sub

    Private Shared Sub SetDocumentViewer(documentViewer As DocumentViewer)

        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document = New DocumentModel()

        Dim section = New Section(document)
        document.Sections.Add(section)

        Dim paragraph = New Paragraph(document)
        section.Blocks.Add(paragraph)

        Dim run = New Run(document, "Hello World!")
        paragraph.Inlines.Add(run)

        Dim xpsDocument = document.ConvertToXpsDocument(SaveOptions.XpsDefault)
        documentViewer.Document = xpsDocument.GetFixedDocumentSequence()

        ' XpsDocument needs to stay referenced so that DocumentViewer can access additional required resources.
        ' Otherwise, GC will collect/dispose XpsDocument and DocumentViewer will not work.
        xpsDocument = document.ConvertToXpsDocument(SaveOptions.XpsDefault)

        documentViewer.Document = xpsDocument.GetFixedDocumentSequence()

    End Sub

End Class
