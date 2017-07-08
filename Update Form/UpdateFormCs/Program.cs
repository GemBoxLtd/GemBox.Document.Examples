using System;
using GemBox.Document;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        DocumentModel document = DocumentModel.Load("FormFilled.docx");

        // Get a snapshot of all form fields in the document.
        var formData = document.Content.FormFieldsData;

        // Update "FullName" text box field.
        FormTextData fullNameData = (FormTextData)formData["FullName"];
        fullNameData.Value = "Jane Doe";

        // Update "BirthDate" text box field.
        FormTextData birthDateData = (FormTextData)formData["BirthDate"];
        birthDateData.Value = new DateTime(2000, 1, 1);

        // Update "Salary" text box field.
        FormTextData salaryData = (FormTextData)formData["Salary"];
        salaryData.Value = 5432.1;

        // Check "Married" check box field.
        FormCheckBoxData marriedData = (FormCheckBoxData)formData["Married"];
        marriedData.Value = true;

        // Select "Female" from drop down field.
        FormDropDownData genderData = (FormDropDownData)formData["Gender"];
        genderData.SelectedItemIndex = genderData.Items.IndexOf("Female");

        document.Save("Update Form.docx");
    }
}
