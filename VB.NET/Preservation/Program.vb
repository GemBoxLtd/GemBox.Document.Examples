Imports System.IO
Imports GemBox.Document

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        ' Load Word document with preservation feature enabled.
        Dim loadOptions As New DocxLoadOptions() With {.PreserveUnsupportedFeatures = True}
        Dim document = DocumentModel.Load("Macros.docm", loadOptions)

        ' Save Word document to output file of same format together with
        ' preserved information (unsupported features) from input file.
        Dim extension As String = Path.GetExtension("Macros.docm")
        document.Save($"Preserved Output{extension}")

    End Sub
End Module