Imports GemBox.Document

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        ' Load input MHTML file.
        Dim document As DocumentModel = DocumentModel.Load("Input.mhtml")

        ' Save output PDF file.
        document.Save("Output.pdf")

    End Sub

End Module
