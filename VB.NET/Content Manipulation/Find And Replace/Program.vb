Imports System.Linq
Imports GemBox.Document

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document = DocumentModel.Load("Reading.docx")

        ' 1. The easiest way how to find and replace text.
        document.Content.Replace(".NET", "C# / VB.NET")

        ' 2. You can also find and highlight text by specifying the format of replacement text.
        Dim highlightText = "read, write, convert and print"
        document.Content.Replace(highlightText, highlightText)
        document.Content.Replace(highlightText, highlightText,
            New CharacterFormat() With {.HighlightColor = Color.Yellow})

        ' 3. You can also search for specified text and achieve more complex replacements.
        ' Notice the "Reverse" method usage for avoiding any possible invalid state due to replacements inside iteration.
        Dim searchText = "GemBox.Document"
        For Each searchedContent As ContentRange In document.Content.Find(searchText).Reverse()

            Dim replaceText = "Word library from GemBox called "
            Dim replacedContent As ContentRange = searchedContent.LoadText(replaceText,
                New CharacterFormat() With {.FontColor = New Color(237, 125, 49)})

            Dim hyperlink As New Hyperlink(document, "https://www.gemboxsoftware.com/document", "GemBox.Document")
            replacedContent.End.InsertRange(hyperlink.Content)

        Next

        document.Save("Find and Replace.docx")

    End Sub
End Module