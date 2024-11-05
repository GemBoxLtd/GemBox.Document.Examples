Imports GemBox.Document
Imports System
Imports System.Linq

Module Program

    Sub Main()
        Example1()
        Example2()
        Example3()
    End Sub

    Sub Example1()
        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document = DocumentModel.Load("ManipulateContent.docx")
        Dim section = document.Sections(0)

        ' Set content of 1st paragraph using plain text.
        section.Blocks(0).Content.LoadText("Inserted plain text to first paragraph.")

        ' Set content of 2nd paragraph using hyperlink.
        Dim hyperlink As New Hyperlink(document, "https://www.gemboxsoftware.com/", "Inserted hyperlink.")
        section.Blocks(1).Content.Set(hyperlink.Content)

        ' Insert HTML text at the end of 3rd paragraph.
        section.Blocks(2).Content.End _
            .LoadText("<p style='color:orange'>Inserted HTML text with orange color.</p>",
                New HtmlLoadOptions() With {.InheritCharacterFormat = True, .InheritParagraphFormat = True})

        ' Insert picture at the beginning of last paragraph.
        Dim picture As New Picture(document, "Dices.png", 40, 30)
        section.Blocks.Last().Content.Start.InsertRange(picture.Content)

        document.Save("InsertContent.docx")
    End Sub

    Sub Example2()
        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document = DocumentModel.Load("ManipulateContent.docx")
        Dim section = document.Sections(0)

        ' Get content from 1st paragraph.
        Dim firstParagraphContent As ContentRange = section.Blocks(0).Content
        Console.WriteLine(firstParagraphContent.ToString())

        ' Get content from 2nd and 3rd paragraphs.
        Dim multipleParagraphsContent As New ContentRange(
            section.Blocks(1).Content.Start,
            section.Blocks(2).Content.End)
        Console.WriteLine(multipleParagraphsContent.ToString())
    End Sub

    Sub Example3()
        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document = DocumentModel.Load("ManipulateContent.docx")
        Dim section = document.Sections(0)

        ' Delete content from 1st and 2nd paragraph.
        Dim multipleParagraphsContent As New ContentRange(
            section.Blocks(0).Content.Start,
            section.Blocks(1).Content.End)
        multipleParagraphsContent.Delete()

        ' Delete content from last (4th) paragraph.
        Dim lastParagraphContent As ContentRange = section.Blocks.Last().Content
        lastParagraphContent.Delete()

        document.Save("DeleteContent.docx")
    End Sub

End Module
