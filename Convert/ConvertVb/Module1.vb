Imports System
Imports GemBox.Document

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As DocumentModel = DocumentModel.Load("Reading.docx")

        ' In order to achieve the conversion of a loaded Word file to PDF,
        ' or to some other Word format,
        ' we just need to save a DocumentModel object to desired output file format.

        document.Save("Convert.pdf")

    End Sub

End Module