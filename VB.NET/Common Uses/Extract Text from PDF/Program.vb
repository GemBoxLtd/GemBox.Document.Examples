Imports System
Imports System.Linq
Imports GemBox.Document
Imports GemBox.Document.Tables

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document = DocumentModel.Load("CustomInvoice.pdf")

        ' Display file's properties.
        Dim properties = document.DocumentProperties
        Console.WriteLine($"Title: {properties.BuiltIn(BuiltInDocumentProperty.Title)}")
        Console.WriteLine($"Author: {properties.BuiltIn(BuiltInDocumentProperty.Author)}")
        Console.WriteLine()

        ' Get paragraphs.
        Dim paragraphs = document.GetChildElements(True, ElementType.Paragraph).Cast(Of Paragraph)()

        ' Get tables.
        Dim tables = document.GetChildElements(True, ElementType.Table).Cast(Of Table)()

        ' Display paragraphs and tables count.
        Console.WriteLine($"Paragraph count: {paragraphs.Count()}")
        Console.WriteLine($"Table count: {tables.Count()}")
        Console.WriteLine()

        ' Display first paragraph's content.
        Dim paragraph = paragraphs.First()
        Console.WriteLine("Paragraph content:")
        Console.WriteLine(paragraph.Content.ToString())

        ' Display last table's content.
        Dim table = tables.Last()
        Console.WriteLine("Table content:")

        For Each row In table.Rows
            Console.WriteLine(New String("-"c, 56))
            For Each cell In row.Cells
                Console.Write($"{cell.Content.ToString().TrimEnd().PadRight(13)}|")
            Next
            Console.WriteLine()
        Next

    End Sub
End Module