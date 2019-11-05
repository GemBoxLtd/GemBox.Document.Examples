Imports System
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
        Dim document = DocumentModel.Load("MergePictures.docx")

        ' Using picture's path, stream and byte array as a data source.
        Dim dataSource = New With
        {
            .Img1 = "Penguins.png",
            .Img2 = File.OpenRead("Penguins.png"),
            .Img3 = File.ReadAllBytes("Penguins.png")
        }

        document.MailMerge.Execute(dataSource)

        document.Save("Merge Pictures.docx")
    End Sub

    Sub Example2()
        Dim document = DocumentModel.Load("MergePicturesWithTemplates.docx")

        Dim dataSource = New With
        {
            .Pic1 = "Penguins.png",
            .Pic2 = "Penguins.png",
            .Pic3 = "Penguins.png"
        }

        ' You can use FieldMerging event to set picture properties
        ' that were not specified in template placeholder.
        AddHandler document.MailMerge.FieldMerging,
            Sub(sender, e)
                If TypeOf e.Inline Is Picture Then
                    DirectCast(e.Inline, Picture).Metadata.Description = "Three dancing penguins."
                End If
            End Sub

        document.MailMerge.Execute(dataSource)

        document.Save("Merge Pictures With Templates.pdf")
    End Sub

End Module