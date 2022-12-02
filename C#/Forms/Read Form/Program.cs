using System;
using GemBox.Document;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = DocumentModel.Load("FormFilled.docx");

        // Get a snapshot of all form fields in the document.
        FormFieldDataCollection formFieldsData = document.Content.FormFieldsData;

        Console.WriteLine($" {"Field type",-20} | {"Name",-20} | {"Value",-20} | {"Value type",-20} ");
        Console.WriteLine(new string('-', 88));

        // Read type, name, value and value type of each form field in the document.
        foreach (FormFieldData formFieldData in formFieldsData)
        {
            Type fieldType = formFieldData.GetType();
            string fieldName = formFieldData.Name;
            object fieldValue = formFieldData.Value;
            Type valueType = fieldValue.GetType();

            Console.WriteLine($" {fieldType.Name,-20} | {fieldName,-20} | {fieldValue,-20} | {valueType.FullName,-20} ");
        }
    }
}