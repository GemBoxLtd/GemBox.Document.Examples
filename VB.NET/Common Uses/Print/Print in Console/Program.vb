Imports GemBox.Document

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As DocumentModel = DocumentModel.Load("Print.docx")

        ' Set Word document's page options.
        For Each section As Section In document.Sections

            Dim pageSetup As PageSetup = section.PageSetup
            pageSetup.Orientation = Orientation.Landscape
            pageSetup.LineNumberRestartSetting = LineNumberRestartSetting.NewPage
            pageSetup.LineNumberDistanceFromText = 50

            Dim pageMargins As PageMargins = pageSetup.PageMargins
            pageMargins.Top = 20
            pageMargins.Left = 100

        Next

        ' Print Word document to default printer (e.g. 'Microsoft Print to Pdf').
        Dim printerName As String = Nothing
        document.Print(printerName)

    End Sub
End Module
