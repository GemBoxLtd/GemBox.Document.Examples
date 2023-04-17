Imports System.Linq
Imports GemBox.Document
Imports GemBox.Document.MailMerging

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        ' If using the Professional version, remove this FreeLimitReached event handler.
        AddHandler ComponentInfo.FreeLimitReached, Sub(sender, e) e.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial

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