Imports System.IO
Imports System.IO.Compression
Imports GemBox.Document

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        ' If using the Professional version, remove this FreeLimitReached event handler.
        AddHandler ComponentInfo.FreeLimitReached, Sub(sender, e) e.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial

        Example1()
        Example2()
        Example3()

    End Sub

    Sub Example1()
        ' Load a Word file into the DocumentModel object.
        Dim document = DocumentModel.Load("Input.docx")

        ' Create image save options.
        Dim imageOptions As New ImageSaveOptions(ImageSaveFormat.Png) With
        {
            .PageNumber = 0, ' Select the first Word page.
            .Width = 1240 ' Set the image width and keep the aspect ratio.
        }

        ' Save the DocumentModel object to a PNG file.
        document.Save("Output.png", imageOptions)
    End Sub

    Sub Example2()
        ' Load a Word file.
        Dim document = DocumentModel.Load("Input.docx")

        ' Max integer value indicates that all document pages should be saved.
        Dim imageOptions As New ImageSaveOptions(ImageSaveFormat.Tiff) With
        {
            .PageCount = Integer.MaxValue
        }

        ' Save the TIFF file with multiple frames, each frame represents a single Word page.
        document.Save("Output.tiff", imageOptions)
    End Sub

    Sub Example3()
        ' Load a Word file.
        Dim document = DocumentModel.Load("Input.docx")

        Dim imageOptions As New ImageSaveOptions()

        ' Get Word pages.
        Dim pages = document.GetPaginator().Pages

        ' Create a ZIP file for storing PNG files.
        Using archiveStream = File.OpenWrite("Output.zip")
            Using archive As New ZipArchive(archiveStream, ZipArchiveMode.Create)
                ' Iterate through the Word pages.
                For pageIndex As Integer = 0 To pages.Count - 1

                    Dim page As DocumentModelPage = pages(pageIndex)

                    ' Create a ZIP entry for each document page.
                    Dim entry = archive.CreateEntry($"Page {pageIndex + 1}.png")

                    ' Save each document page as a PNG image to the ZIP entry.
                    Using imageStream As New MemoryStream()
                        Using entryStream = entry.Open()
                            page.Save(imageStream, imageOptions)
                            imageStream.CopyTo(entryStream)
                        End Using
                    End Using
                Next
            End Using
        End Using
    End Sub

End Module