Imports System
Imports System.Globalization
Imports System.Text
Imports GemBox.Document

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As DocumentModel = DocumentModel.Load("FormFilled.docx")

        ' Get a snapshot of all form fields in the document.
        Dim formFieldsData = document.Content.FormFieldsData

        Dim sb = New StringBuilder()

        ' Write header.
        sb.AppendLine("Document contains following form fields:")
        sb.AppendFormat(CultureInfo.InvariantCulture,
            "{0,-16}|{1,20} = {2,-20}|({3})",
            "Type",
            """"c + "Name" + """"c,
            "Value",
            "Value type").
            AppendLine().AppendLine(New String("-"c, 78))

        ' Write type, name, value and value type of each form field in the document.
        For Each formFieldData In formFieldsData
            sb.AppendFormat(CultureInfo.InvariantCulture,
                "{0,-16}|{1,20} = {2,-20}|({3})",
                formFieldData.GetType().Name,
                """"c + formFieldData.Name + """"c,
                formFieldData.Value,
                If(formFieldData.Value IsNot Nothing, formFieldData.Value.GetType().FullName, "null")).
                AppendLine()
        Next

        Console.WriteLine(sb.ToString())

    End Sub

End Module