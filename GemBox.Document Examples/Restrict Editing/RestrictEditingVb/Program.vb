Imports GemBox.Document

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document = DocumentModel.Load("Reading.docx")

        Dim restriction As EditingRestrictionType = EditingRestrictionType.NoChanges
        Dim password As String = "pass"
        document.Protection.StartEnforcingProtection(restriction, password)

        document.Save("Restrict Editing.docx")

    End Sub
End Module