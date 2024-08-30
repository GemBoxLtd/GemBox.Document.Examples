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

        Dim document As DocumentModel = DocumentModel.Load("Invoice.docx")

        ' Get content from each paragraph.
        For Each paragraph As Paragraph In document.GetChildElements(True, ElementType.Paragraph)
            Console.WriteLine($"Paragraph: {paragraph.Content.ToString()}")
        Next

        ' Get content from each bold run.
        For Each run As Run In document.GetChildElements(True, ElementType.Run)
            If run.CharacterFormat.Bold Then
                Console.WriteLine($"Bold run: {run.Content.ToString()}")
            End If
        Next
    End Sub

    Sub Example2()
        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document = DocumentModel.Load("Reading.docx")

        ' Delete 1st paragraph's inlines.
        Dim paragraph1 = document.Sections(0).Blocks.Cast(Of Paragraph)(0)
        paragraph1.Inlines.Content.Delete()

        ' Delete 3rd and 4th run from the 2nd paragraph.
        Dim paragraph2 = document.Sections(0).Blocks.Cast(Of Paragraph)(1)
        Dim runsContent = New ContentRange(
            paragraph2.Inlines(2).Content.Start,
            paragraph2.Inlines(3).Content.End)
        runsContent.Delete()

        ' Delete specified text content.
        Dim bracketContent = document.Content.Find("(").First()
        bracketContent.Delete()

        document.Save("Delete Content.docx")
    End Sub

    Sub Example3()
        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As New DocumentModel()

        ' Create the whole document using fluent API.
        document.Content.Start _
            .LoadText("First paragraph.") _
            .InsertRange(New Paragraph(document, "Second paragraph.").Content) _
            .LoadText(vbLf) _
            .LoadText("Paragraph with bold text.", New CharacterFormat() With {.Bold = True})

        Dim section = document.Sections(0)

        ' Prepend text to second paragraph.
        section.Blocks(1).Content.Start.LoadText(" Some Prefix ", New CharacterFormat() With {.Subscript = True})

        ' Append text to second paragraph.
        section.Blocks(1).Content.End.LoadText(" Some Suffix ", New CharacterFormat() With {.Superscript = True})

        ' Insert HTML paragraph before third paragraph.
        section.Blocks(2).Content.Start.LoadText("<p style='font:italic 11pt Calibri;color:royalblue;'>Paragraph from HTML content with blue and italic text.</p>",
            New HtmlLoadOptions())

        ' Insert RTF paragraph after fourth paragraph.
        section.Blocks(3).Content.End.LoadText("{\rtf1\ansi\deff0{\colortbl ;\red255\green128\blue64;}\cf1 Paragraph from RTF content with orange text.\par\par}",
            New RtfLoadOptions())

        document.Save("Insert Content.docx")
    End Sub

End Module
