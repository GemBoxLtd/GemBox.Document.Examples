Imports System
Imports GemBox.Document

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As DocumentModel = DocumentModel.Load("MergeFields.docx")

        Dim customer = New With {.CustomerName = "John", .Surname = "Doe", .Date = DateTime.Now}

        document.MailMerge.Execute(customer)

        document.Save("Merge Fields.docx")

    End Sub

End Module