Imports System
Imports GemBox.Document

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As DocumentModel = DocumentModel.Load("Reading.docx")
        Dim documentProperties = document.DocumentProperties

        Console.WriteLine("# Built-in document properties:")

        ' Read built-in document properties.
        For Each builtinProperty In documentProperties.BuiltIn
            Console.WriteLine($"{builtinProperty.Key,25}: {builtinProperty.Value}")
        Next

        ' Write custom document properties.
        documentProperties.Custom.Add("My Custom Property 1", "My Custom Value")
        documentProperties.Custom.Add("My Custom Property 2", 123.4)

        Console.WriteLine()
        Console.WriteLine("# Custom document properties:")

        ' Read custom document properties.
        For Each customProperty In documentProperties.Custom
            Console.WriteLine($"{customProperty.Key,25}: {customProperty.Value} [{customProperty.Value.GetType()}]")
        Next

    End Sub
End Module