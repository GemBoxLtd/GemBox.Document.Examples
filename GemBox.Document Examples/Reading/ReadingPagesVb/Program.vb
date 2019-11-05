Imports System
Imports System.Windows
Imports GemBox.Document

Module Program

    <STAThread>
    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document = DocumentModel.Load("Reading.docx")
        Dim pages = document.GetPaginator().Pages
        Dim count = pages.Count

        For i = 0 To count - 1
            Console.WriteLine(New String("-"c, 50))
            Console.WriteLine($"Page {i + 1} of {count}")
            Console.WriteLine(New String("-"c, 50))

            ' Get FrameworkElement object from Word document's page.
            Dim page As DocumentModelPage = pages(i)
            Dim pageContent As FrameworkElement = pages(i).PageContent

            ' Extract text from FrameworkElement object.
            Console.WriteLine(pageContent.ToText())
        Next

    End Sub
End Module
