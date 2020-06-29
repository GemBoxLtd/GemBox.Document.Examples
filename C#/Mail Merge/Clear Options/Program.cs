using GemBox.Document;
using GemBox.Document.MailMerging;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = DocumentModel.Load("MergeClearOptions.docx");

        // Data source with "Populated" value, but no "Empty" value.
        var data = new { Populated = "sample value" };

        // Execute mail merge on "Example1" merge range.
        // Also, remove fields and paragraphs that didn't merge in this range.
        document.MailMerge.ClearOptions = MailMergeClearOptions.RemoveUnusedFields | MailMergeClearOptions.RemoveEmptyParagraphs;
        document.MailMerge.Execute(data, "Example1");

        // Execute mail merge on "Example2" merge range.
        // Also, remove rows that didn't merge in this range.
        document.MailMerge.ClearOptions = MailMergeClearOptions.RemoveEmptyTableRows;
        document.MailMerge.Execute(data, "Example2");

        // Execute mail merge on "Example3" merge range.
        // Also, remove fields and tables that didn't merge in this range.
        document.MailMerge.ClearOptions = MailMergeClearOptions.RemoveUnusedFields | MailMergeClearOptions.RemoveEmptyTables;
        document.MailMerge.Execute(data, "Example3");

        // Execute mail merge on "Example4" merge range.
        // Also, remove the range if all fields didn't merge in this range.
        document.MailMerge.ClearOptions = MailMergeClearOptions.RemoveEmptyRanges;
        document.MailMerge.Execute(data, "Example4");

        document.Save("Merged Clear Options Output.docx");
    }
}