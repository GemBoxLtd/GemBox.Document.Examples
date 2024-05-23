Imports GemBox.Document

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As New DocumentModel()

        ' Create document with labels and form fields.
        Dim fullNameField As New Field(document, FieldType.FormText, Nothing, New String(ChrW(&H2002), 5))
        fullNameField.FormData.Name = "FullName"

        Dim birthDateField As New Field(document, FieldType.FormText, Nothing, "dd/mm/yyyy")
        birthDateField.FormData.Name = "BirthDate"

        Dim marriedField As New Field(document, FieldType.FormCheckBox)
        marriedField.FormData.Name = "Married"

        Dim genderField As New Field(document, FieldType.FormDropDown)
        genderField.FormData.Name = "Gender"

        document.Sections.Add(
            New Section(document,
                New Paragraph(document, New Run(document, "Full name: "), fullNameField),
                New Paragraph(document, New Run(document, "Birth date: "), birthDateField),
                New Paragraph(document, New Run(document, "Married: "), marriedField),
                New Paragraph(document, New Run(document, "Gender: "), genderField)))

        ' Customize form fields.
        Dim formFieldsData = document.Content.FormFieldsData

        Dim fullNameFieldData = DirectCast(formFieldsData("FullName"), FormTextData)
        fullNameFieldData.MaximumLength = 50
        fullNameFieldData.ValueFormat = "Title case"

        Dim birthdateFieldData = DirectCast(formFieldsData("BirthDate"), FormTextData)
        birthdateFieldData.TextType = FormTextType.Date
        birthdateFieldData.ValueFormat = "dd/MM/yyyy"

        Dim marriedFieldData = DirectCast(formFieldsData("Married"), FormCheckBoxData)
        ' Status text is shown on status bar, a bottom-right corner of Microsoft Word.
        marriedFieldData.StatusText = "Check if you're married."
        ' Help text is shown on field when in focus and F1 is pressed.
        marriedFieldData.HelpText = "Check if you're married."

        Dim genderFieldData = DirectCast(formFieldsData("Gender"), FormDropDownData)
        ' First item as a default option for non-selected value.
        genderFieldData.Items.Add("<Select Gender>")
        genderFieldData.Items.Add("Male")
        genderFieldData.Items.Add("Female")

        ' Make document a form by restricting editing to filling-in form fields.
        document.Protection.StartEnforcingProtection(EditingRestrictionType.FillingForms, "pass")

        document.Save("Create Form.docx")

    End Sub
End Module
