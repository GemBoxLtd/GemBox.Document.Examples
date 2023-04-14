Imports System.Linq
Imports GemBox.Document
Imports GemBox.Document.MailMerging

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("DN-2023Apr13-XtlswQ30Ua/FqbsDIZzCR9HjIB0FccANeA9uJUm8Fu8xKLo//xMsyDrDFztBh0CqKEvYarPkDGYCenzpGznopWYJvHg==A")

        Dim numberOfLabels As Integer = 10
        Dim document = DocumentModel.Load("MergeLabels.docx")

        ' Create data source with multiple records.
        Dim source = Enumerable.Range(1, numberOfLabels).Select(
            Function(i) New With {.Name = "Person " & i, .Company = "Company " & i})

        ' Specify clear options to remove unmerged fields.
        document.MailMerge.ClearOptions = MailMergeClearOptions.RemoveUnusedFields

        ' Execute mail merge process.
        document.MailMerge.Execute(source)

        document.Save("MergeLabelsOutput.docx")

    End Sub
End Module