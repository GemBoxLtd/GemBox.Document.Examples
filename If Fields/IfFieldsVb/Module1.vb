Imports System
Imports GemBox.Document

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As DocumentModel = DocumentModel.Load("MergeIfFields.docx")

        Dim customer = New With {.Gender = "M", .CustomerName = "John", .Surname = "Doe"}

        document.MailMerge.Execute(customer)

        document.Save("If Fields.docx")

    End Sub

End Module