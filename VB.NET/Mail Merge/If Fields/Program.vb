Imports GemBox.Document

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document = DocumentModel.Load("MergeIfFields.docx")

        Dim data = New With
        {
            .FirstName = "John",
            .LastName = "Doe",
            .Gender = "Male",
            .Age = 30
        }

        document.MailMerge.Execute(data)

        document.Save("Merged If Fields Output.docx")

    End Sub
End Module