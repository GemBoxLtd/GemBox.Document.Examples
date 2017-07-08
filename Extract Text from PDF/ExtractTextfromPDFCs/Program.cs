using System;
using System.Text;
using System.Text.RegularExpressions;
using GemBox.Document;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        DocumentModel document = DocumentModel.Load("CustomInvoice.pdf");

        StringBuilder sb = new StringBuilder();

        // Read PDF file's document properties.
        sb.AppendFormat("Author: {0}", document.DocumentProperties.BuiltIn[BuiltInDocumentProperty.Author]).AppendLine();
        sb.AppendFormat("DateContentCreated: {0}", document.DocumentProperties.BuiltIn[BuiltInDocumentProperty.DateLastSaved]).AppendLine();

        // Sample's input parameter.
        string pattern = @"(?<WorkHours>\d+)\s+(?<UnitPrice>\d+\.\d{2})\s+(?<Total>\d+\.\d{2})";
        Regex regex = new Regex(pattern);

        int row = 0;
        StringBuilder line = new StringBuilder();

        // Read PDF file's text content and match a specified regular expression.
        foreach (Match match in regex.Matches(document.Content.ToString()))
        {
            line.Length = 0;
            line.AppendFormat("Result: {0}: ", ++row);

            // Either write only successfully matched named groups or entire match.
            bool hasAny = false;
            for (int i = 0; i < match.Groups.Count; ++i)
            {
                string groupName = regex.GroupNameFromNumber(i);
                Group matchGroup = match.Groups[i];
                if (matchGroup.Success && groupName != i.ToString())
                {
                    line.AppendFormat("{0}= {1}, ", groupName, matchGroup.Value);
                    hasAny = true;
                }
            }

            if (hasAny)
                line.Length -= 2;
            else
                line.Append(match.Value);

            sb.AppendLine(line.ToString());
        }

        Console.WriteLine(sb.ToString());
    }
}
