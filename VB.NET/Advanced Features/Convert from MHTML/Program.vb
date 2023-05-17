Imports GemBox.Document

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        ' If using the Professional version, remove this FreeLimitReached event handler.
        AddHandler ComponentInfo.FreeLimitReached, Sub(sender, e) e.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial

        ' Load input MHTML file.
        Dim document As DocumentModel = DocumentModel.Load("Input.mhtml")

        ' Save output PDF file.
        document.Save("Output.pdf")

    End Sub
End Module
