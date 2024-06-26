Imports GemBox.Document
Imports System.IO
Imports System.IO.Compression

Module Program

    Sub Main()
        Example1()
        Example2()
        Example3()
    End Sub

    Sub Example1()
        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        ' In order to convert Word to PDF, we just need to:
        '   1. Load DOC or DOCX file into DocumentModel object.
        '   2. Save DocumentModel object to PDF file.
        Dim document As DocumentModel = DocumentModel.Load("Input.docx")
        document.Save("Output.pdf")
    End Sub

    Sub Example2()
        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        ' Load Word file.
        Dim document As DocumentModel = DocumentModel.Load("Input.docx")

        ' Get Word pages.
        Dim pages = document.GetPaginator().Pages

        ' Create PDF save options.
        Dim pdfSaveOptions As New PdfSaveOptions() With {.ImageDpi = 220}

        ' Create ZIP file for storing PDF files.
        Using archiveStream = File.OpenWrite("Output.zip")
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
        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim conformanceLevel As PdfConformanceLevel = PdfConformanceLevel.PdfA1a

        ' Load Word file.
        Dim document As DocumentModel = DocumentModel.Load("Input.docx")

        ' Create PDF save options.
        Dim options As New PdfSaveOptions() With
        {
            .ConformanceLevel = conformanceLevel
        }

        ' Save to PDF file.
        document.Save("OutputWithConformance.pdf", options)
    End Sub

End Module
