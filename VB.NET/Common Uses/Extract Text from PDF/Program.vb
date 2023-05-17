Imports System
Imports System.Linq
Imports GemBox.Document
Imports GemBox.Document.Tables

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        ' If using the Professional version, remove this FreeLimitReached event handler.
        AddHandler ComponentInfo.FreeLimitReached, Sub(sender, e) e.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial

        Dim document = DocumentModel.Load("CustomInvoice.pdf")

        ' Get paragraphs.
        Dim paragraphs = document.GetChildElements(True, ElementType.Paragraph).Cast(Of Paragraph)()

        ' Get tables.
        Dim tables = document.GetChildElements(True, ElementType.Table).Cast(Of Table)()

        ' Display paragraphs and tables count.
        Console.WriteLine($"Paragraph count: {paragraphs.Count()}")
        Console.WriteLine($"Table count: {tables.Count()}")
        Console.WriteLine()

        ' Display first paragraph's content.
        Dim paragraph = paragraphs.FirstOrDefault()
        If paragraph IsNot Nothing Then
            Console.WriteLine("Paragraph content:")
            Console.WriteLine(paragraph.Content.ToString())
        End If

        ' Display last table's content.
        Dim table = tables.LastOrDefault()
        If paragraph IsNot Nothing Then
            Console.WriteLine("Table content:")
            For Each row In table.Rows
                For Each cell In row.Cells
                    Console.Write($"{cell.Content.ToString().TrimEnd().PadRight(15)}|")
                Next
                Console.WriteLine()
            Next
        End If

    End Sub
End Module