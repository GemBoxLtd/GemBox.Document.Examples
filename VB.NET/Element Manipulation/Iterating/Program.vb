Imports System
Imports System.Linq
Imports GemBox.Document

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As DocumentModel = DocumentModel.Load("Reading.docx")

        Dim numberOfSections As Integer = document.Sections.Count
        Dim numberOfParagraphs As Integer = document.GetChildElements(True, ElementType.Paragraph).Count()
        Dim numberOfRunsAndFields As Integer = document.GetChildElements(True, ElementType.Run, ElementType.Field).Count()

        Dim section = document.Sections(0)

        Dim numberOfElements As Integer = section.GetChildElements(True).Count()
        Dim numberOfBlocks As Integer = section.GetChildElements(True).OfType(Of Block)().Count()
        Dim numberOfInlines As Integer = section.GetChildElements(True).OfType(Of Inline)().Count()

        Console.WriteLine("File has:")
        Console.WriteLine($" - {numberOfSections} sections.")
        Console.WriteLine($" - {numberOfParagraphs} paragraphs.")
        Console.WriteLine($" - {numberOfRunsAndFields} runs and fields.")

        Console.WriteLine()

        Console.WriteLine("First section has:")
        Console.WriteLine($" - {numberOfElements} elements.")
        Console.WriteLine($" - {numberOfBlocks} blocks.")
        Console.WriteLine($" - {numberOfInlines} inlines.")

    End Sub
End Module