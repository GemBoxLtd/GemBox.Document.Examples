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
        var formData = document.Content.FormFieldsData;

        // Update "FullName" text field.
        var fullNameData = (FormTextData)formData["FullName"];
        fullNameData.Value = "Jane Doe";

        // Update "BirthDate" text field.
        var birthDateData = (FormTextData)formData["BirthDate"];
        birthDateData.Value = new DateTime(2000, 2, 29);

        // Check "Married" check-box field.
        var marriedData = (FormCheckBoxData)formData["Married"];
        marriedData.Value = true;

        // Select "Female" from drop-down field.
        var genderData = (FormDropDownData)formData["Gender"];
        genderData.SelectedItemIndex = genderData.Items.IndexOf("Female");

        document.Save("Form Updated.docx");
    }
}