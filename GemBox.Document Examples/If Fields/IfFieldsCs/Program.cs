using GemBox.Document;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = DocumentModel.Load("MergeIfFields.docx");

        var data = new
        {
            FirstName = "John",
            LastName = "Doe",
            Gender = "Male",
            Age = 30
        };

        document.MailMerge.Execute(data);

        document.Save("Merged If Fields Output.docx");
    }
}