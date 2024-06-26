using GemBox.Document;
using System;
using System.Linq;

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

        var document = DocumentModel.Load("Reading.docx");
        var documentProperties = document.DocumentProperties;

        Console.WriteLine("# Built-in document properties:");

        // Write built-in document properties.
        documentProperties.BuiltIn[BuiltInDocumentProperty.Title] = "My Title";
        documentProperties.BuiltIn[BuiltInDocumentProperty.DateLastSaved] = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ssZ");

        // Read built-in document properties.
        foreach (var builtinProperty in documentProperties.BuiltIn)
            Console.WriteLine($"{builtinProperty.Key,20}: {builtinProperty.Value}");

        Console.WriteLine();
        Console.WriteLine("# Custom document properties:");

        // Write custom document properties.
        documentProperties.Custom["My Custom Property 1"] = "My Custom Value";
        documentProperties.Custom["My Custom Property 2"] = 123.4;

        // Read custom document properties.
        foreach (var customProperty in documentProperties.Custom)
            Console.WriteLine($"{customProperty.Key,20}: {customProperty.Value,-20} [{customProperty.Value.GetType()}]");
    }

    static void Example2()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = DocumentModel.Load("PropertiesAndVariables.docx");

        // Update built-in and custom document properties.
        document.DocumentProperties.BuiltIn[BuiltInDocumentProperty.Title] = "Updated Title";
        document.DocumentProperties.Custom["Custom1"] = "Updated custom value 1!";
        document.DocumentProperties.Custom["Custom2"] = "Updated custom value 2!";

        // Update document variables.
        document.Variables["Variable1"] = "Updated variable value 1!";
        document.Variables["Variable2"] = "Updated variable value 2!";

        var fields = document.GetChildElements(true, ElementType.Field)
            .OfType<Field>()
            .Where(f => f.FieldType == FieldType.DocProperty || f.FieldType == FieldType.DocVariable);

        // Update DOCPROPERTY and DOCVARIABLE fields.
        foreach (var field in fields)
            field.Update();

        document.Save("UpdatedPropertiesAndVariables.docx");
    }
}
