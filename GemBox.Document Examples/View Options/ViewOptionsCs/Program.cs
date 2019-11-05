using GemBox.Document;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");
        
        var document = new DocumentModel();

        document.Sections.Add(
            new Section(document,
                new Paragraph(document, " ")));

        var viewOptions = document.ViewOptions;
        viewOptions.ViewType = ViewType.Web;
        viewOptions.Zoom = 75;

        document.Save("View Options.docx");
    }
}