using System.Linq;
using GemBox.Document;
using GemBox.Document.MailMerging;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        int numberOfLabels = 10;
        var document = DocumentModel.Load("MergeLabels.docx");

        // Create data source with multiple records.
        var source = Enumerable.Range(1, numberOfLabels).Select(
            i => new { Name = "Person " + i, Company = "Company " + i });

        // Specify clear options to remove unmerged fields.
        document.MailMerge.ClearOptions = MailMergeClearOptions.RemoveUnusedFields;

        // Execute mail merge process.
        document.MailMerge.Execute(source);

        document.Save("MergeLabelsOutput.docx");
    }
}