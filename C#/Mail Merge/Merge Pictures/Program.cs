using System;
using System.IO;
using GemBox.Document;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        Example1();
        Example2();
    }

    static void Example1()
    {
        var document = DocumentModel.Load("MergePictures.docx");

        // Using picture's path, stream and byte array as a data source.
        var dataSource = new
        {
            Img1 = "Penguins.png",
            Img2 = File.OpenRead("Penguins.png"),
            Img3 = File.ReadAllBytes("Penguins.png"),
        };

        document.MailMerge.Execute(dataSource);

        document.Save("Merge Pictures.docx");
    }

    static void Example2()
    {
        var document = DocumentModel.Load("MergePicturesWithTemplates.docx");

        var dataSource = new
        {
            Pic1 = "Penguins.png",
            Pic2 = "Penguins.png",
            Pic3 = "Penguins.png"
        };

        // You can use FieldMerging event to set picture properties
        // that were not specified in template placeholder.
        document.MailMerge.FieldMerging += (sender, e) =>
        {
            if (e.Inline is Picture picture)
                picture.Metadata.Description = "Three dancing penguins.";
        };

        document.MailMerge.Execute(dataSource);

        document.Save("Merge Pictures With Templates.pdf");
    }
}