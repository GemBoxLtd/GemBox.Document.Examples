Imports GemBox.Document
Imports System.Linq

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document = DocumentModel.Load("Reading.docx")

        ' Get first Section element.
        Dim section = document.Sections(0)

        ' Get first Paragraph element.
        Dim paragraph = section.Blocks.OfType(Of Paragraph).First()

        ' Clone paragraph and add it to section.
        Dim cloneParagraph = paragraph.Clone(True)
        section.Blocks.Add(cloneParagraph)

        ' Clone section and add it to document.
        Dim cloneSection = section.Clone(True)
        document.Sections.Add(cloneSection)

        document.Save("Cloning.docx")

    End Sub
End Module
