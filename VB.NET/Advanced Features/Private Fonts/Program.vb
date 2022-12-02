Imports GemBox.Document

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        FontSettings.FontsBaseDirectory = "."

        Dim document As New DocumentModel()

        document.DefaultCharacterFormat.FontName = "Almonte Snow"
        document.DefaultCharacterFormat.Size = 48

        document.Content.LoadText("Hello World!")

        document.Save("Private Fonts.pdf")

    End Sub
End Module