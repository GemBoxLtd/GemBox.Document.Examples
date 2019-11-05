Imports System
Imports System.IO
Imports System.IO.Compression
Imports GemBox.Document

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Example1()
        Example2()
        Example3()
        Example4()

    End Sub

    Sub Example1()
        ' In order to convert Word to PDF, we just need to:
        ' 1. Load DOC or DOCX file into DocumentModel object.
        ' 2. Save DocumentModel object to PDF file.
        Dim document As DocumentModel = DocumentModel.Load("Input.docx")
        document.Save("Output1.pdf")
    End Sub

    Sub Example2()
        ' Load Word file.
        Dim document As DocumentModel = DocumentModel.Load("Input.docx")

        ' Get Word pages.
        Dim pages = document.GetPaginator().Pages

        ' Create PDF save options.
        Dim pdfSaveOptions As New PdfSaveOptions() With {.ImageDpi = 220}

        ' Create ZIP file for storing PDF files.
        Using archiveStream = File.OpenWrite("Output2.zip")
            Using archive As New ZipArchive(archiveStream, ZipArchiveMode.Create)
                ' Iterate through Word pages.
                For pageIndex As Integer = 0 To pages.Count - 1

                    Dim page As DocumentModelPage = pages(pageIndex)

                    ' Create ZIP entry for each document page.
                    Dim entry = archive.CreateEntry($"Page {pageIndex + 1}.pdf")

                    ' Save each document page as PDF to ZIP entry.
                    Using pdfStream As New MemoryStream()
                        Using entryStream = entry.Open()
                            page.Save(pdfStream, pdfSaveOptions)
                            pdfStream.CopyTo(entryStream)
                        End Using
                    End Using
                Next
            End Using
        End Using
    End Sub

    Sub Example3()
        ' Load input HTML file.
        Dim document As DocumentModel = DocumentModel.Load("Input.html")

        ' When reading any HTML content a single Section element is created.
        ' We can use that Section element to specify various page options.
        Dim section As Section = document.Sections(0)
        Dim pageSetup As PageSetup = section.PageSetup
        Dim pageMargins As PageMargins = pageSetup.PageMargins
        With pageMargins
            .Left = 0
            .Right = 0
            .Top = 0
            .Bottom = 0
        End With

        ' Save output PDF file.
        document.Save("Output3.pdf")
    End Sub

    Sub Example4()
        Dim html = "
<html>
<style>
  @page {
    size: A5 landscape;
    margin: 6cm 1cm 1cm;
    mso-header-margin: 1cm;
    mso-footer-margin: 1cm;
  }

  body {
    background: #EDEDED;
    border: 1pt solid black;
    padding: 20pt;
  }

  br {
    page-break-before: always;
  }

  p { margin: 0; }
  header { color: #FF0000; text-align: center; }
  main { color: #00B050; }
  footer { color: #0070C0; text-align: right; }
</style>

<body>
  <header>
    <p>Header text.</p>
  </header>
  <main>
    <p>First page.</p>
    <br>
    <p>Second page.</p>
    <br>
    <p>Third page.</p>
    <br>
    <p>Fourth page.</p>
  </main>
  <footer>
    <p>Footer text.</p>
    <p>Page <span style='mso-field-code:PAGE'>1</span> of <span style='mso-field-code:NUMPAGES'>1</span></p>
  </footer>
</body>
</html>"

        Dim htmlLoadOptions As New HtmlLoadOptions()
        Using htmlStream As New MemoryStream(htmlLoadOptions.Encoding.GetBytes(html))

            ' Load input HTML text as stream.
            Dim document = DocumentModel.Load(htmlStream, htmlLoadOptions)
            ' Save output PDF file.
            document.Save("Output4.pdf")

        End Using
    End Sub

End Module