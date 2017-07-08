Imports System
Imports GemBox.Document

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As DocumentModel = New DocumentModel

        ' Set the content for the whole document
        document.Content.LoadText("Paragraph 1" & vbLf & "Paragraph 2" & vbLf & "Paragraph 3" & vbLf & "Paragraph 4" & vbLf & "Paragraph 5")

        Dim bold = New CharacterFormat() With {
            .Bold = True
        }

        ' Set the content for the 2nd paragraph
        document.Sections(0).Blocks(1).Content.LoadText("Bold paragraph 2", bold)

        ' Set the content for 3rd and 4th paragraph to be the same as the content of 1st and 2nd paragraph
        Dim para3 = document.Sections(0).Blocks(2)
        Dim para4 = document.Sections(0).Blocks(3)
        Dim destinationRange = New ContentRange(para3.Content.Start, para4.Content.End)
        Dim para1 = document.Sections(0).Blocks(0)
        Dim para2 = document.Sections(0).Blocks(1)
        Dim sourceRange = New ContentRange(para1.Content.Start, para2.Content.End)
        destinationRange.Set(sourceRange)

        ' Set content using HTML tags
        document.Sections(0).Blocks(4).Content.LoadText("Paragraph 5 <b>(part of this paragraph is bold)</b>", LoadOptions.HtmlDefault)

        document.Save("Set Content.docx")

    End Sub

End Module