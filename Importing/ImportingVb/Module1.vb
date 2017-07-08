Imports System
Imports System.IO
Imports GemBox.Document

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As DocumentModel = DocumentModel.Load("Invoice.docx")

        Dim pathToFileDirectory As String = "Resources"

        Dim sourceDocument As DocumentModel = DocumentModel.Load(Path.Combine(pathToFileDirectory, "Reading.docx"), LoadOptions.DocxDefault)

        ' Reuse same mapping for importing all sections to improve performance.
        Dim mapping = New ImportMapping(sourceDocument, document, False)

        ' Import all sections from source document.
        For Each sourceSection As Section In sourceDocument.Sections
            Dim destinationSection As Section = document.Import(sourceSection, True, mapping)
            document.Sections.Add(destinationSection)
        Next

        document.Save("Importing.docx")

    End Sub

End Module