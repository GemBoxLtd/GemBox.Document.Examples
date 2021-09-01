Imports System
Imports System.Text.RegularExpressions
Imports GemBox.Document

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As DocumentModel = DocumentModel.Load("CustomInvoice.pdf")
        Dim properties As DocumentProperties = document.DocumentProperties

        ' Read PDF file's properties.
        Console.WriteLine($"Author: {properties.BuiltIn(BuiltInDocumentProperty.Author)}")
        Console.WriteLine($"Created on: {properties.BuiltIn(BuiltInDocumentProperty.DateContentCreated)}")
        Console.WriteLine()

        ' Read PDF file's text content and match specified regular expression.
        Dim text = document.Content.ToString()
        Dim regex As New Regex("(?<Hours>\d+)\s+(?<Unit>\d+\.\d{2})\s+(?<Price>\d+\.\d{2})")

        For Each match As Match In regex.Matches(text)
            Dim groups = match.Groups
            Console.WriteLine($"Hours={groups("Hours")} | Unit={groups("Unit")} | Price={groups("Price")}")
        Next

    End Sub
End Module