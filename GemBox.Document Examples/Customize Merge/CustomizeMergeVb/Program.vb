Imports GemBox.Document

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
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
End Module