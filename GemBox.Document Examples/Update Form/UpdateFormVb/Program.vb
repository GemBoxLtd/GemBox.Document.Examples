Imports System
Imports GemBox.Document

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document = DocumentModel.Load("FormFilled.docx")

        ' Get a snapshot of all form fields in the document.
        Dim formData = document.Content.FormFieldsData

        ' Update "FullName" text field.
        Dim fullNameData = DirectCast(formData("FullName"), FormTextData)
        fullNameData.Value = "Jane Doe"

        ' Update "BirthDate" text field.
        Dim birthDateData = DirectCast(formData("BirthDate"), FormTextData)
        birthDateData.Value = New DateTime(2000, 2, 29)

        ' Check "Married" check-box field.
        Dim marriedData = DirectCast(formData("Married"), FormCheckBoxData)
        marriedData.Value = True

        ' Select "Female" from drop-down field.
        Dim genderData = DirectCast(formData("Gender"), FormDropDownData)
        genderData.SelectedItemIndex = genderData.Items.IndexOf("Female")

        document.Save("Form Updated.docx")

    End Sub
End Module