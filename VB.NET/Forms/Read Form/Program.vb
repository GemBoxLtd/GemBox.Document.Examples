Imports GemBox.Document
Imports System

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document = DocumentModel.Load("FormFilled.docx")

        ' Get a snapshot of all form fields in the document.
        Dim formFieldsData As FormFieldDataCollection = document.Content.FormFieldsData

        Console.WriteLine($" {"Field type",-20} | {"Name",-20} | {"Value",-20} | {"Value type",-20} ")
        Console.WriteLine(New String("-"c, 88))

        ' Read type, name, value and value type of each form field in the document.
        For Each formFieldData As FormFieldData In formFieldsData

            Dim fieldType As Type = formFieldData.GetType()
            Dim fieldName As String = formFieldData.Name
            Dim fieldValue As Object = formFieldData.Value
            Dim valueType As Type = fieldValue.GetType()

            Console.WriteLine($" {fieldType.Name,-20} | {fieldName,-20} | {fieldValue,-20} | {valueType.FullName,-20} ")

        Next

    End Sub
End Module
