Imports GemBox.Document
Imports System
Imports System.Linq

Module Program

    Sub Main()

        Example1()
        Example2()

    End Sub

    Sub Example1()
        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As DocumentModel = DocumentModel.Load("Reading.docx")
        Dim documentProperties = document.DocumentProperties

        Console.WriteLine("# Built-in document properties:")

        ' Write built-in document properties.
        documentProperties.BuiltIn(BuiltInDocumentProperty.Title) = "My Title"
        documentProperties.BuiltIn(BuiltInDocumentProperty.DateLastSaved) = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ssZ")

        ' Read built-in document properties.
        For Each builtinProperty In documentProperties.BuiltIn
            Console.WriteLine($"{builtinProperty.Key,20}: {builtinProperty.Value}")
        Next

        Console.WriteLine()
        Console.WriteLine("# Custom document properties:")

        ' Write custom document properties.
        documentProperties.Custom("My Custom Property 1") = "My Custom Value"
        documentProperties.Custom("My Custom Property 2") = 123.4

        ' Read custom document properties.
        For Each customProperty In documentProperties.Custom
            Console.WriteLine($"{customProperty.Key,20}: {customProperty.Value,-20} [{customProperty.Value.GetType()}]")
        Next
    End Sub

    Sub Example2()
        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document = DocumentModel.Load("PropertiesAndVariables.docx")

        ' Update built-in and custom document properties.
        document.DocumentProperties.BuiltIn(BuiltInDocumentProperty.Title) = "Updated Title"
        document.DocumentProperties.Custom("Custom1") = "Updated custom value 1!"
        document.DocumentProperties.Custom("Custom2") = "Updated custom value 2!"

        ' Update document variables.
        document.Variables("Variable1") = "Updated variable value 1!"
        document.Variables("Variable2") = "Updated variable value 2!"

        Dim fields = document.GetChildElements(True, ElementType.Field) _
            .OfType(Of Field)() _
            .Where(Function(f) f.FieldType = FieldType.DocProperty OrElse f.FieldType = FieldType.DocVariable)
        
        ' Update DOCPROPERTY and DOCVARIABLE fields.
        For Each field In fields
            field.Update()
        Next

        document.Save("UpdatedPropertiesAndVariables.docx")
    End Sub

End Module
