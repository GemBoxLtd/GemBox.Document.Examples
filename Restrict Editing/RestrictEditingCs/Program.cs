using System;
using GemBox.Document;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = DocumentModel.Load("Reading.docx");

        // Disallow all editing in the document (document is read-only).
        // Since password is not specified, all users can stop enforcing protection in MS Word.
        document.Protection.StartEnforcingProtection(EditingRestrictionType.NoChanges, null);

        document.Save("Restrict Editing.docx");
    }
}
