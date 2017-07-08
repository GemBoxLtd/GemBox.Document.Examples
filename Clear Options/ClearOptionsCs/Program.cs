using System;
using GemBox.Document;
using GemBox.Document.MailMerging;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        DocumentModel document = DocumentModel.Load("MailMergeClearOptions.docx");

        // Example 1: Data source will remove "First choice" paragraph because there is no value defined for FirstChoice field.
        var dataSource1 = new
        {
            Header = "My header",
            SecondChoice = "I have chosen second choice."
        };

        document.MailMerge.ClearOptions = MailMergeClearOptions.RemoveEmptyParagraphs;
        document.MailMerge.Execute(dataSource1, "Example1");

        // Example 2: Data source will remove table row with label "Address" because value for field Address is null.
        var dataSource2 = new
        {
            Name = "John Doe",
            Email = "john.doe@acme.com",
            Address = (string)null,
            Country = "USA"
        };

        document.MailMerge.ClearOptions = MailMergeClearOptions.RemoveEmptyTableRows;
        document.MailMerge.Execute(dataSource2, "Example2");

        // Example 3: Data source will remove mail merge range for second item because it has both Header and Content values that are string.Empty or null.
        var dataSource3 = new
        {
            Count = 2,
            HeaderContent = new object[]
            {
                new
                {
                    Header = "My header 1",
                    Content = "My content 1.\nSecond line of my content 1."
                },
                new
                {
                    Header = string.Empty,
                    Content = (object)null
                },
                new
                {
                    Header = "My header 3",
                    Content = "My content 3.\nSecond line of my content 3."
                }
            }
        };
        document.MailMerge.ClearOptions = MailMergeClearOptions.RemoveEmptyRanges | MailMergeClearOptions.RemoveEmptyParagraphs;
        document.MailMerge.Execute(dataSource3, "Example3");

        document.Save("Clear Options.docx");
    }
}
