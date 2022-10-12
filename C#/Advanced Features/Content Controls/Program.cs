using System;
using System.Linq;
using System.Text;
using GemBox.Document;
using GemBox.Document.CustomMarkups;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        Example1();
        Example2();
        Example3();
    }

    static void Example1()
    {
        var document = new DocumentModel();

        var section = new Section(document);
        document.Sections.Add(section);

        // Create locked Rich Text Content Control.
        var richTextControl = new BlockContentControl(document, ContentControlType.RichText,
            new Paragraph(document, "This text is inside Rich Text Content Control."),
            new Paragraph(document, "It cannot be deleted or edited."));
        richTextControl.Properties.LockEditing = true;
        richTextControl.Properties.LockDeleting = true;
        section.Blocks.Add(richTextControl);

        // Create named Plain Text Content Control.
        var plainTextControl = new BlockContentControl(document, ContentControlType.PlainText,
            new Paragraph(document, "Plain Text Content Control with tag and title."));
        plainTextControl.Properties.Tag = "Plain Text Name";
        plainTextControl.Properties.Title = "Plain Text Title";
        section.Blocks.Add(plainTextControl);

        // Create CheckBox Content Control.
        var checkBoxControl = new InlineContentControl(document, ContentControlType.CheckBox,
           new Run(document, "☒") { CharacterFormat = { FontName = "MS Gothic" } });
        checkBoxControl.Properties.Checked = true;

        // Create ComboBox Content Control.
        var comboBoxControl = new InlineContentControl(document, ContentControlType.ComboBox,
            new Run(document, "<Select GemBox Component>"));
        comboBoxControl.Properties.ListItems.Add(new ContentControlListItem("<Select GemBox Component>", "NONE"));
        comboBoxControl.Properties.ListItems.Add(new ContentControlListItem("GemBox.Spreadsheet", "GBS"));
        comboBoxControl.Properties.ListItems.Add(new ContentControlListItem("GemBox.Document", "GBD"));
        comboBoxControl.Properties.ListItems.Add(new ContentControlListItem("GemBox.Pdf", "GBA"));
        comboBoxControl.Properties.ListItems.Add(new ContentControlListItem("GemBox.Presentation", "GBP"));
        comboBoxControl.Properties.ListItems.Add(new ContentControlListItem("GemBox.Email", "GBE"));

        section.Blocks.Add(new Paragraph(document,
            checkBoxControl,
            new SpecialCharacter(document, SpecialCharacterType.LineBreak),
            comboBoxControl));

        document.Save("Content Controls.docx");
    }

    static void Example2()
    {
        var document = DocumentModel.Load("XmlMapping.docx");

        // Edit mapped XML part.
        document.CustomXmlParts[0].Data = Encoding.UTF8.GetBytes(
@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<customer>
    <firstName>Jane</firstName>
    <lastName>Doe</lastName>
    <birthday>2010-01-01T00:00:00</birthday>
    <married>true</married>
</customer>");

        // Update Content Controls inlines or blocks based on the values from mapped XML part.
        foreach (var contentControl in document.GetChildElements(true).OfType<IContentControl>())
            contentControl.Update();

        document.Save("Updated ContentControls.docx");
    }

    static void Example3()
    {
        var document = DocumentModel.Load("XmlMapping.docx");

        // Edit Content Controls.
        foreach (var contentControl in document.GetChildElements(true).OfType<IContentControl>())
        {
            switch (contentControl.Properties.Title)
            {
                case "FirstName":
                    contentControl.Content.LoadText("Joe");
                    break;
                case "LastName":
                    contentControl.Content.LoadText("Smith");
                    break;
                case "Birthday":
                    contentControl.Properties.Date = new DateTime(2002, 2, 2);
                    break;
                case "Married":
                    contentControl.Properties.Checked = true;
                    break;
            }

            // Update mapped XML part based on the content from Content Control.
            contentControl.UpdateSource();
        }

        document.Save("Updated XmlMapping.docx");
    }
}