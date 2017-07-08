Imports System
Imports GemBox.Document

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim inputPassword As String = "inpass"
        Dim outputPassword As String = "outpass"
        Dim ownerPassword As String = ""

        Dim document = DocumentModel.Load("PdfEncryption.pdf", New PdfLoadOptions With {.Password = inputPassword})
        Dim options = New PdfSaveOptions() With
        {
            .DocumentOpenPassword = outputPassword,
            .PermissionsPassword = ownerPassword,
            .Permissions = PdfPermissions.None
        }

        document.Save("PDF Encryption.pdf", options)

    End Sub

End Module