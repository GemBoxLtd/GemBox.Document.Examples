Imports GemBox.Document

Module Program

    Sub Main()

        Example1()
        Example2()

    End Sub

    Sub Example1()
        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document = DocumentModel.Load("MergeCustomizations.docx")

        AddHandler document.MailMerge.FieldMerging,
            Sub(sender, e)
                If e.IsValueFound And e.Value IsNot Nothing Then
                    Select Case e.FieldName

                        Case "CheckedField"
                            Dim checkedValue As Boolean = DirectCast(e.Value, Boolean)
                            Dim run = DirectCast(e.Inline, Run)
                            run.CharacterFormat.FontColor = If(checkedValue, Color.Green, Color.Red)
                            run.Text = If(checkedValue, "☑", "☒")
                            Exit Select

                        Case "LinkField"
                            Dim linkValue = DirectCast(e.Value, (Address As String, DisplayText As String))
                            e.Inline = New Hyperlink(e.Document, linkValue.Address, linkValue.DisplayText)
                            Exit Select

                        Case "ImageField"
                            Dim imagePath = e.Value.ToString()
                            e.Inline = New Picture(e.Document, imagePath)
                            Exit Select

                    End Select
                End If
            End Sub

        document.MailMerge.Execute(
            New With
            {
                .CheckedField = True,
                .LinkField = (Address:="https://www.gemboxsoftware.com/", DisplayText:="GemBox Homepage"),
                .ImageField = "Dices.png"
            })

        document.Save("Merged Customizations Output.docx")
    End Sub

    Sub Example2()
        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document = DocumentModel.Load("MergeCustomTypes.docx")

        ' Map merge fields with prefix to data source without prefix.
        ' E.g. "Html:MyName" field's name to "MyName" data source name.
        For Each fieldName As String In document.MailMerge.GetMergeFieldNames()
            Dim index As Integer = fieldName.IndexOf(":"c)
            If index > 0 AndAlso Not document.MailMerge.FieldMappings.ContainsKey(fieldName) Then
                document.MailMerge.FieldMappings.Add(fieldName, fieldName.Substring(index + 1))
            End If
        Next

        ' Customize mail merge to support our custom prefixes.
        AddHandler document.MailMerge.FieldMerging,
            Sub(sender, e)
                If Not e.IsValueFound Then Return

                Dim customImport As Boolean = True

                If e.FieldName.StartsWith("Html:") Then
                    e.Field.Content.End.LoadText(CStr(e.Value), LoadOptions.HtmlDefault)
                ElseIf e.FieldName.StartsWith("Rtf:") Then
                    e.Field.Content.End.LoadText(CStr(e.Value), LoadOptions.RtfDefault)
                ElseIf e.FieldName.StartsWith("Docx:") Then
                    e.Field.Content.End.InsertRange(DocumentModel.Load(CStr(e.Value)).Sections(0).Blocks.Content)
                Else
                    customImport = False
                End If

                If customImport Then
                    ' Remove the default import.
                    e.Inline = Nothing

                    ' Check if the content is inserted into the same parent paragraph or added after it.
                    ' This depends on whether the data source contains inline-level or block-level elements.
                    ' If it's added after, then remove the empty parent which used to contain merge field.
                    If e.Field.ParentCollection.Count = 1 Then e.Field.Parent.Content.Delete()
                End If
            End Sub

        document.MailMerge.Execute(
            New With
            {
                .Field1 = "{\rtf1\ansi\deff0{\fonttbl{\f0 Arial Black;}}{\colortbl ;\red255\green128\blue64;}\f0\cf1 This is rich formatted text (in RTF format).}",
                .Field2 = "<p style='font-family:Arial Narrow;color:royalblue;'>This is another rich formatted text (in HTML format).</p>",
                .Field3 = "<p style='font-family:Arial Narrow;color:seagreen;'>And another rich formatted text (in HTML format).</p>",
                .Field4 = "Reading.docx"
            })

        document.Save("Merged Custom Types Output.docx")
    End Sub

End Module
