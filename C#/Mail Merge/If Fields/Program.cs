using GemBox.Document;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        Example1();
        Example2();
    }

    static void Example1()
    {
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

    static void Example2()
    {
        var document = DocumentModel.Load("MergeConditionalFields.docx");

        var data = new
        {
            FirstName = "Jane",
            LastName = "Doe",
            Gender = "Female",
            Age = 30,
            Married = true
        };

        document.MailMerge.Execute(data);

        document.Save("Merged Complex Conditional Fields Output.docx");
    }
}