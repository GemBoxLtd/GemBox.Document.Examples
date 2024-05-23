Imports GemBox.Document

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim userPassword As String = "pass"
        Dim ownerPassword As String = ""
        Dim permissions As PdfPermissions = PdfPermissions.None

        Dim document = DocumentModel.Load("Reading.docx")

        Dim options As New PdfSaveOptions() With
        {
            .DocumentOpenPassword = userPassword,
            .PermissionsPassword = ownerPassword,
            .Permissions = permissions
        }

        document.Save("PDF Encryption.pdf", options)

    End Sub
End Module
