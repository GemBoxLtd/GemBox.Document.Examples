Imports GemBox.Document
Imports GemBox.Document.CustomMarkups
Imports System
Imports System.Linq
Imports System.Text

Module Program

    Sub Main()

        Example1()
        Example2()
        Example3()

    End Sub

    Sub Example1()
        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As New DocumentModel()

        Dim section As New Section(document)
        document.Sections.Add(section)

        ' Create locked Rich Text Content Control.
        Dim richTextControl As New BlockContentControl(document, ContentControlType.RichText,
            New Paragraph(document, "This text is inside Rich Text Content Control."),
            New Paragraph(document, "It cannot be deleted or edited."))
        richTextControl.Properties.LockEditing = True
        richTextControl.Properties.LockDeleting = True
        section.Blocks.Add(richTextControl)

        ' Create named Plain Text Content Control.
        Dim plainTextControl As New BlockContentControl(document, ContentControlType.PlainText,
            New Paragraph(document, "Plain Text Content Control with tag and title."))
        plainTextControl.Properties.Tag = "Plain Text Name"
        plainTextControl.Properties.Title = "Plain Text Title"
        section.Blocks.Add(plainTextControl)

        ' Create CheckBox Content Control.
        Dim checkBoxControl As New InlineContentControl(document, ContentControlType.CheckBox)
        checkBoxControl.Properties.CharacterFormat = New CharacterFormat() With {.FontName = "MS Gothic"}
        checkBoxControl.Properties.Checked = True

        ' Create ComboBox Content Control.
        Dim comboBoxControl As New InlineContentControl(document, ContentControlType.ComboBox,
            New Run(document, "<Select GemBox Component>"))
        comboBoxControl.Properties.ListItems.Add(New ContentControlListItem("<Select GemBox Component>", "NONE"))
        comboBoxControl.Properties.ListItems.Add(New ContentControlListItem("GemBox.Spreadsheet", "GBS"))
        comboBoxControl.Properties.ListItems.Add(New ContentControlListItem("GemBox.Document", "GBD"))
        comboBoxControl.Properties.ListItems.Add(New ContentControlListItem("GemBox.Pdf", "GBA"))
        comboBoxControl.Properties.ListItems.Add(New ContentControlListItem("GemBox.Presentation", "GBP"))
        comboBoxControl.Properties.ListItems.Add(New ContentControlListItem("GemBox.Email", "GBE"))
        comboBoxControl.Properties.ListItems.Add(New ContentControlListItem("GemBox.Imaging", "GBI"))

        section.Blocks.Add(New Paragraph(document,
            checkBoxControl,
            New SpecialCharacter(document, SpecialCharacterType.LineBreak),
            comboBoxControl))

        document.Save("Content Controls.docx")
    End Sub

    Sub Example2()
        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document = DocumentModel.Load("XmlMapping.docx")

        ' Edit mapped XML part.
        document.CustomXmlParts(0).Data = Encoding.UTF8.GetBytes(
"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<customer>
    <firstName>Jane</firstName>
    <lastName>Doe</lastName>
    <birthday>2010-01-01T00:00:00</birthday>
    <married>true</married>
</customer>")

        ' Update Content Controls inlines or blocks based on the values from mapped XML part.
        For Each contentControl In document.GetChildElements(True).OfType(Of IContentControl)()
            contentControl.Update()
        Next

        document.Save("Updated ContentControls.docx")
    End Sub

    Sub Example3()
        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document = DocumentModel.Load("XmlMapping.docx")

        ' Edit Content Controls.
        For Each contentControl In document.GetChildElements(True).OfType(Of IContentControl)()
            Select Case contentControl.Properties.Title
                Case "FirstName"
                    contentControl.Content.LoadText("Joe")
                Case "LastName"
                    contentControl.Content.LoadText("Smith")
                Case "Birthday"
                    contentControl.Properties.Date = New DateTime(2002, 2, 2)
                Case "Married"
                    contentControl.Properties.Checked = True
            End Select

            ' Update mapped XML part based on the content from Content Control.
            contentControl.UpdateSource()
        Next

        document.Save("Updated XmlMapping.docx")
    End Sub

End Module
