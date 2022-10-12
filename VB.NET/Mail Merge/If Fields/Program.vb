Imports GemBox.Document

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Example1()
        Example2()

    End Sub

    Sub Example1()
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

    Sub Example2()
        Dim document = DocumentModel.Load("MergeConditionalFields.docx")

        Dim data = New With
        {
            .FirstName = "Jane",
            .LastName = "Doe",
            .Gender = "Female",
            .Age = 30,
            .Married = True
        }

        document.MailMerge.Execute(data)

        document.Save("Merged Complex Conditional Fields Output.docx")
    End Sub
End Module