using GemBox.Document;
using GemBox.Document.Vba;

class Program
{
    static void Main()
    {
        Example1();
        Example2();
    }

    static void Example1()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = new DocumentModel();

        var section = new Section(document);
        document.Sections.Add(section);

        // Create the module.
        VbaModule vbaModule = document.VbaProject.Modules.Add("SampleModule", VbaModuleType.Document);
        vbaModule.Code = 
@"Sub WriteHello()
    Selection.TypeText Text:=""Hello World!""
End Sub";

        // Save the document as macro-enabled Word file.
        document.Save("AddVbaModule.docm");
    }

    static void Example2()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = DocumentModel.Load("SampleVba.docm");

        // Get the module.
        VbaModule vbaModule = document.VbaProject.Modules["ThisDocument"];
        // Update text for the popup message.
        vbaModule.Code = vbaModule.Code.Replace("Hello world!", "Hello from GemBox.Document!");

        document.Save("UpdateVbaModule.docm");
    }
}
