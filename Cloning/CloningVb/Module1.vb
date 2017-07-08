Imports System
Imports System.Linq
Imports GemBox.Document

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As DocumentModel = DocumentModel.Load("Reading.docx")

        ' Clone section
        document.Sections.Add(document.Sections(0).Clone(True))

        ' Clone paragraphs
        For Each item As Block In document.Sections(0).Blocks
            document.Sections.Last().Blocks.Add(item.Clone(True))
        Next

        document.Save("Cloning.docx")

    End Sub

End Module