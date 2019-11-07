Imports System.IO
Imports System.Xml
Imports System.Linq
Imports GemBox.Document
Imports GemBox.Document.CustomMarkups

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Example1()
        Example2()

    End Sub

    Sub Example1()
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
        Dim checkBoxControl As New InlineContentControl(document, ContentControlType.CheckBox,
           New Run(document, "☒") With {.CharacterFormat = New CharacterFormat() With {.FontName = "MS Gothic"}})
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

        section.Blocks.Add(New Paragraph(document,
            checkBoxControl,
            New SpecialCharacter(document, SpecialCharacterType.LineBreak),
            comboBoxControl))

        document.Save("Content Controls.docx")
    End Sub

    Sub Example2()
        Dim document = DocumentModel.Load("XmlMapping.docx")

        ' Get the Content Control.
        Dim contentControl = CType(document.GetChildElements(True, ElementType.InlineContentControl).First(), InlineContentControl)
        Dim xmlMapping = contentControl.Properties.XmlMapping

        ' Get the mapped XML part.
        Dim xmlPart = xmlMapping.CustomXmlPart

        ' Create XmlDocument from XML.
        Dim xmlDocument = New XmlDocument()
        xmlDocument.Load(New MemoryStream(xmlPart.Data))

        ' Locate the node to which is the Content Control mapped using XPath.
        Dim node = xmlDocument.SelectSingleNode(xmlMapping.XPath)

        ' Change the node value.
        node.InnerText = "Jonathan"

        ' Update the XmlPart Data.
        Dim outputMemoryStream = New MemoryStream()
        xmlDocument.Save(outputMemoryStream)
        xmlPart.Data = outputMemoryStream.ToArray()

        ' Get the node value.
        Dim nodeValue = node.InnerText

        ' Update Content Control inlines.
        contentControl.Inlines.Clear()
        contentControl.Inlines.Add(New Run(document, nodeValue) With
        {
            .CharacterFormat = contentControl.Properties.CharacterFormat
        })

        document.Save("Xml Mapping.docx")
    End Sub

End Module