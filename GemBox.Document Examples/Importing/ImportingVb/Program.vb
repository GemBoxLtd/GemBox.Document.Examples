Imports GemBox.Document

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim destinationDocument = DocumentModel.Load("Invoice.docx")
        Dim sourceDocument = DocumentModel.Load("Reading.docx")

        ' 1. The following is the easiest way how to import content from source to destination,
        ' it will automatically imports the document's elements.
        'destinationDocument.Content.End.InsertRange(sourceDocument.Content);

        ' 2. The following enables various customization.
        ' For instance, we can change that the imported content doesn't start from new page
        ' by setting "SectionStart.Continuous" on first destination section.

        ' Reuse the same mapping for importing to improve performance.
        Dim mapping As New ImportMapping(sourceDocument, destinationDocument, False)

        ' Import all sections from source document to destination document.
        For Each sourceSection In sourceDocument.Sections

            Dim destinationSection = destinationDocument.Import(sourceSection, True, mapping)
            destinationDocument.Sections.Add(destinationSection)

        Next

        destinationDocument.Save("Importing.docx")

    End Sub
End Module