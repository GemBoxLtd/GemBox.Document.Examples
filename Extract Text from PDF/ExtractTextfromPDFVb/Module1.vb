Imports System
Imports System.Text
Imports System.Text.RegularExpressions
Imports GemBox.Document

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As DocumentModel = DocumentModel.Load("CustomInvoice.pdf")

        Dim sb As New StringBuilder()

        ' Read PDF file's document properties.
        sb.AppendFormat("Author: {0}", document.DocumentProperties.BuiltIn(BuiltInDocumentProperty.Author)).AppendLine()
        sb.AppendFormat("DateContentCreated: {0}", document.DocumentProperties.BuiltIn(BuiltInDocumentProperty.DateLastSaved)).AppendLine()

        ' Sample's input parameter.
        Dim pattern As String = "(?<WorkHours>\d+)\s+(?<UnitPrice>\d+\.\d{2})\s+(?<Total>\d+\.\d{2})"
        Dim regex As Regex = New Regex(pattern)

        Dim row As Integer = 0
        Dim line As StringBuilder = New StringBuilder()

        ' Read PDF file's text content and match a specified regular expression.
        For Each match As Match In regex.Matches(document.Content.ToString())
            line.Length = 0
            line.AppendFormat("Result: {0}: ", ++row)

            ' Either write only successfully matched named groups or entire match.
            Dim hasAny As Boolean = False
            For i As Integer = 1 To match.Groups.Count - 1
                Dim groupName As String = regex.GroupNameFromNumber(i)
                Dim matchGroup As Group = match.Groups(i)
                If (matchGroup.Success And groupName <> i.ToString()) Then
                    line.AppendFormat("{0}= {1}, ", groupName, matchGroup.Value)
                    hasAny = True
                End If
            Next

            If (hasAny) Then
                line.Length -= 2
            Else
                line.Append(match.Value)
            End If

            sb.AppendLine(line.ToString())
        Next

        Console.WriteLine(sb.ToString())

    End Sub

End Module