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

        StringBuilder sb = new StringBuilder();

        foreach (Paragraph paragraph in document.GetChildElements(true, ElementType.Paragraph))
        {
            foreach (Run run in paragraph.GetChildElements(true, ElementType.Run))
            {
                bool isBold = run.CharacterFormat.Bold;
                string text = run.Text;

                sb.AppendFormat("{0}{1}{2}", isBold ? "<b>" : "", text, isBold ? "</b>" : "");
            }
            sb.AppendLine();
        }

        Console.WriteLine(sb.ToString());
    }
}
