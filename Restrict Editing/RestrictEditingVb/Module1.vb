Imports System
Imports GemBox.Document

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document = DocumentModel.Load("Reading.docx")

        ' Disallow all editing in the document (document is read-only).
        ' Since password is not specified, all users can stop enforcing protection in MS Word.
        document.Protection.StartEnforcingProtection(EditingRestrictionType.NoChanges, Nothing)

        document.Save("Restrict Editing.docx")

    End Sub

End Module