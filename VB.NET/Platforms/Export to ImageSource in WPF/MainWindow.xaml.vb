Imports System.Windows
Imports GemBox.Document

Namespace ExportToImageSource
    Partial Public Class MainWindow
        Inherits Window

        Public Sub New()

            Me.InitializeComponent()

            ' If using the Professional version, put your serial key below.
            ComponentInfo.SetLicense("FREE-LIMITED-KEY")

            Me.SetImageControl("Reading.docx", 0)

        End Sub

        Private Sub SetImageControl(path As String, pageIndex As Integer)

            Dim document = DocumentModel.Load(path)

            Dim imageOptions As New ImageSaveOptions()
            imageOptions.PageNumber = pageIndex

            Dim imageSource = document.ConvertToImageSource(imageOptions)
            Me.ImageControl.Source = imageSource

        End Sub

    End Class
End Namespace