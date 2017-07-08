Imports System.Windows
Imports System.Windows.Controls
Imports GemBox.Document

Class MainWindow

    Public Sub New()
        InitializeComponent()

        SetDocumentViewer(Me.ImageControl)
    End Sub

    Private Shared Sub SetDocumentViewer(image As Image)

        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document = New DocumentModel()

        Dim section = New Section(document)
        document.Sections.Add(section)

        Dim paragraph = New Paragraph(document)
        section.Blocks.Add(paragraph)

        Dim run = New Run(document, "Hello World!")
        paragraph.Inlines.Add(run)

        image.Source = document.ConvertToImageSource(SaveOptions.ImageDefault)

    End Sub

End Class
