using System;
using GemBox.Document;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = DocumentModel.Load("Reading.docx");
        var documentProperties = document.DocumentProperties;

        Console.WriteLine("# Built-in document properties:");

        // Read built-in document properties.
        foreach (var builtinProperty in documentProperties.BuiltIn)
            Console.WriteLine($"{builtinProperty.Key,25}: {builtinProperty.Value}");

        // Write custom document properties.
        documentProperties.Custom.Add("My Custom Property 1", "My Custom Value");
        documentProperties.Custom.Add("My Custom Property 2", 123.4);

        Console.WriteLine();
        Console.WriteLine("# Custom document properties:");

        // Read custom document properties.
        foreach (var customProperty in documentProperties.Custom)
            Console.WriteLine($"{customProperty.Key,25}: {customProperty.Value} [{customProperty.Value.GetType()}]");
    }
}