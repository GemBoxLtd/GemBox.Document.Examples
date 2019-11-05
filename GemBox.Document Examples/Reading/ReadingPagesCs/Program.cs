using System;
using System.Windows;
using GemBox.Document;

class Program
{
    [STAThread]
    static void Main()
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = DocumentModel.Load("Reading.docx");
        var pages = document.GetPaginator().Pages;

        for (int i = 0, count = pages.Count; i < count; ++i)
        {
            Console.WriteLine(new string('-', 50));
            Console.WriteLine($"Page {i + 1} of {count}");
            Console.WriteLine(new string('-', 50));

            // Get FrameworkElement object from Word document's page.
            DocumentModelPage page = pages[i];
            FrameworkElement pageContent = pages[i].PageContent;

            // Extract text from FrameworkElement object.
            Console.WriteLine(pageContent.ToText());
        }
    }
}