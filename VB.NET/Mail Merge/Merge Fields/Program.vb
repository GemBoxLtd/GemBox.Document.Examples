Imports System
Imports GemBox.Document

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document = DocumentModel.Load("MergeFields.docx")

        ' Create data source for mail merge process.
        Dim data = New With
        {
            .Number = 10203,
            .Date = DateTime.Now,
            .Company = "ACME Corp",
            .Address = "240 Old Country Road, Springfield, IL",
            .Country = "USA",
            .FullName = "Joe Smith"
        }

        ' Execute mail merge process.
        document.MailMerge.Execute(data)

        document.Save("Mail Merge Output.docx")

    End Sub
End Module