using System;
using System.Text;
using GemBox.Document;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        DocumentModel document = DocumentModel.Load("Reading.docx");

        var sb = new StringBuilder();

        // Get content from each paragraph
        foreach (Paragraph paragraph in document.GetChildElements(true, ElementType.Paragraph))
        {
            sb.AppendFormat("Paragraph: {0}", paragraph.Content.ToString());
            sb.AppendLine();
        }

        // Get content from each bold run
        foreach (Run run in document.GetChildElements(true, ElementType.Run))
        {
            if (run.CharacterFormat.Bold)
            {
                sb.AppendFormat("Bold run: {0}", run.Content.ToString());
                sb.AppendLine();
            }
        }

        Console.WriteLine(sb.ToString());
    }
}
