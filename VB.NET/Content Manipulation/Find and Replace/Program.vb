Imports System.Linq
Imports System.Text.RegularExpressions
Imports GemBox.Document
Imports GemBox.Document.Tables

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Example1()
        Example2()

    End Sub

    Sub Example1()
        Dim document = DocumentModel.Load("FindAndReplaceText.docx")

        ' The easiest way how you can find and replace text is with "Replace" method.
        document.Content.Replace("%FirstName%", "John")
        document.Content.Replace("%LastName%", "Doe")

        ' Another way would be to use Regex.
        document.Content.Replace(New Regex("%DATE%", RegexOptions.IgnoreCase),
            DateTime.Today.ToLongDateString())

        document.Content.Replace(New Regex("%.*?%"),
            Function(range)
                Dim value As String = Nothing
                Select Case range.ToString()
                    Case "%Address%"
                        value = "240 Old Country Road"
                    Case "%City%"
                        value = "Springfield"
                    Case "%State%"
                        value = "IL"
                    Case "%Country%"
                        value = "USA"
                End Select

                If String.IsNullOrEmpty(value) Then Return range

                Dim format = (CType(range.Start.Parent, Run)).CharacterFormat
                Dim run As New Run(document, value) With {.CharacterFormat = format.Clone()}
                Return run.Content
            End Function)

        ' You can also search for placeholder text with the "Find" method and then achieve a
        ' more complex replacement, like the following which has a replace text with different formatting.
        ' Notice that the "Reverse" extension method is used here to avoid a possible invalid state because
        ' the replacements are done while iterating through the document's content.
        For Each searchedContent As ContentRange In document.Content.Find("%Price%").Reverse()
            Dim replacedContent As ContentRange = searchedContent.LoadText("$",
                New CharacterFormat() With {.Size = 14, .FontColor = Color.Blue, .Bold = True})
            replacedContent.End.LoadText("100.00",
                New CharacterFormat() With {.Size = 11, .FontColor = Color.Purple, .Italic = True})
        Next

        ' Another more complex replacement in which searched text is replaced with a hyperlink.
        For Each searchedContent As ContentRange In document.Content.Find("%Email%").Reverse()
            Dim emailLink As New Hyperlink(document, "mailto:john.doe@example.com", "John.Doe@example.com")
            searchedContent.Set(emailLink.Content)
        Next

        ' You can also find and highlight text by specifying "HighlightColor" of replacement text.
        For Each searchedContent As ContentRange In document.Content.Find("membership").Reverse()
            Dim highlightedText As New Run(document, "membership")
            highlightedText.CharacterFormat = (CType(searchedContent.Start.Parent, Run)).CharacterFormat.Clone()
            highlightedText.CharacterFormat.HighlightColor = Color.Yellow
            searchedContent.Set(highlightedText.Content)
        Next

        document.Save("FoundAndReplacedText.docx")
    End Sub

    Sub Example2()
        Dim document = DocumentModel.Load("FindAndReplaceContent.docx")

        Dim dummyText = "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa."

        ' Find an image placeholder.
        Dim picturePlaceholder = document.Content.Find("%Portrait%").First()
        Dim picture As New Picture(document, "avatar.png")

        ' Replace the placeholder text with the image.
        picturePlaceholder.Set(picture.Content)

        ' Find an HTML placeholder.
        Dim htmlPlaceholder = document.Content.Find("%AboutMe%").First()
        Dim html =
$"<ul style='font:11pt Calibri;'>
    <li style='color:red;'>{dummyText}</li>
    <li style='color:green;'>{dummyText}</li>
    <li style='color:blue;'>{dummyText}</li>
</ul>"

        ' Replace the placeholder text with HTML formatted text.
        htmlPlaceholder.LoadText(html, New HtmlLoadOptions())

        ' Find a table placeholder.
        Dim tablePlaceholder = document.Content.Find("%JobHistory%").First()

        Dim table As New Table(document,
            New TableRow(document,
                New TableCell(document, New Paragraph(document, "2021 - 2030")),
                New TableCell(document, New Paragraph(document, dummyText))),
            New TableRow(document,
                New TableCell(document, New Paragraph(document, "2011 - 2020")),
                New TableCell(document, New Paragraph(document, dummyText))),
            New TableRow(document,
                New TableCell(document, New Paragraph(document, "2001 - 2010")),
                New TableCell(document, New Paragraph(document, dummyText))))

        table.Columns.Add(New TableColumn(70))
        table.Columns.Add(New TableColumn(250))
        table.TableFormat.AutomaticallyResizeToFitContents = False

        ' Delete the placeholder text and insert the table before it.
        tablePlaceholder = tablePlaceholder.LoadText(String.Empty)
        tablePlaceholder.Start.InsertRange(table.Content)

        document.Save("FoundAndReplacedContent.docx")
    End Sub

End Module