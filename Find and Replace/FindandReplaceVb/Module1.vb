Imports System
Imports System.Linq
Imports GemBox.Document

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As DocumentModel = DocumentModel.Load("Reading.docx")

        ' Find and count text.
        Dim documentCount = document.Content.Find("GemBox.Document").Count()

        Dim counter = documentCount

        ' Find text and load another text in its place.
        For Each item As ContentRange In document.Content.Find("GemBox.Document").Reverse()
            item.LoadText(String.Format("GBD ({0}/{1})", counter, documentCount))
            counter -= 1
        Next

        ' Find And replace text.
        document.Content.Replace(".NET", "C# / VB.NET")

        document.Save("Find and Replace.docx")

    End Sub

End Module