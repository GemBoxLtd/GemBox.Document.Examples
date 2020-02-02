Imports GemBox.Document
Imports GemBox.Document.MailMerging

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document = DocumentModel.Load("MergeClearOptions.docx")

        ' Data source with "Populated" value, but no "Empty" value.
        Dim data As New With {.Populated = "sample value"}

        ' Execute mail merge on "Example1" merge range.
        ' Also, remove fields and paragraphs that didn't merge in this range.
        document.MailMerge.ClearOptions = MailMergeClearOptions.RemoveUnusedFields Or MailMergeClearOptions.RemoveEmptyParagraphs
        document.MailMerge.Execute(data, "Example1")

        ' Execute mail merge on "Example2" merge range.
        ' Also, remove rows that didn't merge in this range.
        document.MailMerge.ClearOptions = MailMergeClearOptions.RemoveEmptyTableRows
        document.MailMerge.Execute(data, "Example2")

        ' Execute mail merge on "Example2" merge range.
        ' Also, remove the range if all fields didn't merge in this range.
        document.MailMerge.ClearOptions = MailMergeClearOptions.RemoveEmptyRanges
        document.MailMerge.Execute(data, "Example3")

        document.Save("Merged Clear Options Output.docx")

    End Sub
End Module