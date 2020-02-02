using GemBox.Document;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = DocumentModel.Load("Reading.docx");

        EditingRestrictionType restriction = EditingRestrictionType.NoChanges;
        string password = "pass";
        document.Protection.StartEnforcingProtection(restriction, password);

        document.Save("Restrict Editing.docx");
    }
}