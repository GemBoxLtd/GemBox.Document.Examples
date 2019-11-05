using GemBox.Document;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = new DocumentModel();

        // Create document with labels and form fields.
        document.Sections.Add(
            new Section(document,
                new Paragraph(document,
                    new Run(document, "Full name: "),
                    new Field(document, FieldType.FormText, null, new string('\x2002', 5))
                    { FormData = { Name = "FullName" } }),
                new Paragraph(document,
                    new Run(document, "Birth date: "),
                    new Field(document, FieldType.FormText, null, "dd/mm/yyyy")
                    { FormData = { Name = "BirthDate" } }),
                new Paragraph(document,
                    new Run(document, "Married: "),
                    new Field(document, FieldType.FormCheckBox)
                    { FormData = { Name = "Married" } }),
                new Paragraph(document,
                    new Run(document, "Gender: "),
                    new Field(document, FieldType.FormDropDown)
                    { FormData = { Name = "Gender" } })));

        // Customize form fields.
        var formFieldsData = document.Content.FormFieldsData;

        var fullNameFieldData = (FormTextData)formFieldsData["FullName"];
        fullNameFieldData.MaximumLength = 50;
        fullNameFieldData.ValueFormat = "Title case";

        var birthDateFieldData = (FormTextData)formFieldsData["BirthDate"];
        birthDateFieldData.TextType = FormTextType.Date;
        birthDateFieldData.ValueFormat = "dd/MM/yyyy";

        var marriedFieldData = (FormCheckBoxData)formFieldsData["Married"];
        // Status text is shown on status bar, a bottom-right corner of Microsoft Word.
        marriedFieldData.StatusText = "Check if you're married.";
        // Help text is shown on field when in focus and F1 is pressed.
        marriedFieldData.HelpText = "Check if you're married.";

        var genderFieldData = (FormDropDownData)formFieldsData["Gender"];
        // First item as a default option for non-selected value.
        genderFieldData.Items.Add("<Select Gender>");
        genderFieldData.Items.Add("Male");
        genderFieldData.Items.Add("Female");

        // Make document a form by restricting editing to filling-in form fields.
        document.Protection.StartEnforcingProtection(EditingRestrictionType.FillingForms, "pass");

        document.Save("Create Form.docx");
    }
}