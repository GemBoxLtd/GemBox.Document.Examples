Imports System
Imports GemBox.Document

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As DocumentModel = DocumentModel.Load("FormFilled.docx")

        ' Get a snapshot of all form fields in the document.
        Dim formData = document.Content.FormFieldsData

        ' Update "FullName" text box field.
        Dim fullNameData As FormTextData = DirectCast(formData("FullName"), FormTextData)
        fullNameData.Value = "Jane Doe"

        ' Update "BirthDate" text box field.
        Dim birthDateData As FormTextData = DirectCast(formData("BirthDate"), FormTextData)
        birthDateData.Value = New DateTime(2000, 1, 1)

        ' Update "Salary" text box field.
        Dim salaryData As FormTextData = DirectCast(formData("Salary"), FormTextData)
        salaryData.Value = 5432.1

        ' Check "Married" check box field.
        Dim marriedData As FormCheckBoxData = DirectCast(formData("Married"), FormCheckBoxData)
        marriedData.Value = True

        ' Select "Female" from drop down field.
        Dim genderData As FormDropDownData = DirectCast(formData("Gender"), FormDropDownData)
        genderData.SelectedItemIndex = genderData.Items.IndexOf("Female")

        document.Save("Update Form.docx")

    End Sub

End Module