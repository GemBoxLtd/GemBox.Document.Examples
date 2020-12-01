Imports System.IO
Imports GemBox.Document

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Example1()
        Example2()

    End Sub

    Sub Example1()
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
        document.Save("Output1.pdf")
    End Sub

    Sub Example2()
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
            document.Save("Output2.pdf")

        End Using
    End Sub

End Module