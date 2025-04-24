Imports GemBox.Document

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        ' Load Word document, preservation feature is enabled by default.
        Dim document = DocumentModel.Load("Preservation.docx")

        ' Modify Word document.
        document.Sections(0).Blocks.Insert(0,
            New Paragraph(document, "You can preserve unsupported features when modifying a document!"))

        ' Save Word document to an output file of the same format together with
        ' preserved information (unsupported features) from the input file.
        document.Save("PreservedOutput.docx")

    End Sub

End Module
