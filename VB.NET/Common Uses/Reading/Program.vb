Imports GemBox.Document
Imports System
Imports System.IO
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

        ' Load Word document from file's path.
        Dim document = DocumentModel.Load("Reading.docx")

        ' Get Word document's plain text.
        Dim text As String = document.Content.ToString()

        ' Get Word document's count statistics.
        Dim charactersCount As Integer = text.Replace(Environment.NewLine, String.Empty).Length
        Dim wordsCount As Integer = document.Content.CountWords()
        Dim paragraphsCount As Integer = document.GetChildElements(True, ElementType.Paragraph).Count()
        Dim pageCount As Integer = document.GetPaginator().Pages.Count

        ' Display file's count statistics.
        Console.WriteLine($"Characters count: {charactersCount}")
        Console.WriteLine($"     Words count: {wordsCount}")
        Console.WriteLine($"Paragraphs count: {paragraphsCount}")
        Console.WriteLine($"     Pages count: {pageCount}")
        Console.WriteLine()

        ' Display file's text content.
        Console.WriteLine(text)
    End Sub

    Sub Example2()
        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document = DocumentModel.Load("Reading.docx")
        Using writer = File.CreateText("Output.txt")

            ' Iterate through all Paragraph elements in the Word document.
            For Each paragraph As Paragraph In document.GetChildElements(True, ElementType.Paragraph)

                ' Iterate through all Run elements in the Paragraph element.
                For Each run As Run In paragraph.GetChildElements(True, ElementType.Run)

                    Dim text As String = run.Text
                    Dim format As CharacterFormat = run.CharacterFormat

                    ' Replace text with bold formatting to 'Mathematical Bold Italic' Unicode characters.
                    ' For instance, "ABC" to "𝑨𝑩𝑪".
                    If format.Bold Then
                        text = String.Concat(text.Select(
                            Function(c)
                                Return If(c >= "A"c AndAlso c <= "Z"c, Char.ConvertFromUtf32(119847 + AscW(c)),
                                       If(c >= "a"c AndAlso c <= "z"c, Char.ConvertFromUtf32(119841 + AscW(c)),
                                       c.ToString()))
                            End Function))
                    End If

                    writer.Write(text)
                Next

                writer.WriteLine()
            Next
        End Using
    End Sub

    Sub Example3()
        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document = DocumentModel.Load("Reading.docx")
        Dim pages = document.GetPaginator().Pages

        Dim pageNumber As Integer = 1
        For Each page As DocumentModelPage In pages

            Console.WriteLine(New String("-"c, 50))
            Console.WriteLine($"Page {pageNumber} of {pages.Count}")
            pageNumber += 1
            Console.WriteLine(New String("-"c, 50))

            Using stream As New MemoryStream()
                ' Save Word document's page to TXT file.
                Dim txtSaveOptions = SaveOptions.TxtDefault
                page.Save(stream, txtSaveOptions)

                ' Display page's extracted text.
                Console.WriteLine(txtSaveOptions.Encoding.GetString(stream.ToArray()))
            End Using
        Next
    End Sub

End Module
