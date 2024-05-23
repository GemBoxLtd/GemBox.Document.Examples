Imports GemBox.Document

Module Program

    Sub Main()
        Example1()
        Example2()
    End Sub

    Sub Example1()
        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        ' Load Word file (DOC, DOCX, RTF, XML) into DocumentModel object.
        Dim document = DocumentModel.Load("ExportToHtml.docx")

        Dim saveOptions As New HtmlSaveOptions() With
        {
            .HtmlType = HtmlType.Html,
            .EmbedImages = True,
            .UseSemanticElements = True
        }

        ' Save DocumentModel object to HTML (or MHTML) file.
        document.Save("Exported.html", saveOptions)
    End Sub

    Sub Example2()
        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        ' Load input HTML file.
        Dim document As DocumentModel = DocumentModel.Load("Input.html")

        ' When reading any HTML content a single Section element is created,
        ' which can be used to specify various Word document's page options.
        ' The same can also be achieved with HTML document itself,
        ' by using CSS properties on "@page" directive or "<body>" element.
        Dim section As Section = document.Sections(0)
        Dim pageSetup As PageSetup = section.PageSetup
        Dim pageMargins As PageMargins = pageSetup.PageMargins
        With pageMargins
            .Left = 0
            .Right = 0
            .Top = 0
            .Bottom = 0
        End With

        ' Save output DOCX file.
        document.Save("Output.docx")
    End Sub

End Module
