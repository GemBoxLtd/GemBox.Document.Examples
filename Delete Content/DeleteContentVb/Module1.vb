Imports System
Imports GemBox.Document

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As DocumentModel = DocumentModel.Load("Reading.docx")

        ' Delete paragraph break between 1st and 2nd paragraph (concatenate 1st and 2nd paragraph)
        Dim blocks = document.Sections(0).Blocks
        Dim paragraphBreakRange = New ContentRange(blocks(0).Content.End, blocks(1).Content.Start)
        paragraphBreakRange.Delete()

        ' Delete content of 2nd run
        blocks.Cast(Of Paragraph)(0).Inlines(1).Content.Delete()

        document.Save("Delete Content.docx")

    End Sub

End Module