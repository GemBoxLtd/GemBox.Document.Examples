Imports System.Linq
Imports GemBox.Document

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document = DocumentModel.Load("Reading.docx")

        ' Delete 1st paragraph's inlines.
        Dim paragraph1 = document.Sections(0).Blocks.Cast(Of Paragraph)(0)
        paragraph1.Inlines.Content.Delete()

        ' Delete 3rd and 4th run from the 2nd paragraph.
        Dim paragraph2 = document.Sections(0).Blocks.Cast(Of Paragraph)(1)
        Dim runsContent = New ContentRange(
            paragraph2.Inlines(2).Content.Start,
            paragraph2.Inlines(3).Content.End)
        runsContent.Delete()

        ' Delete specified text content.
        Dim bracketContent = document.Content.Find("(").First()
        bracketContent.Delete()

        document.Save("Delete Content.docx")

    End Sub
End Module