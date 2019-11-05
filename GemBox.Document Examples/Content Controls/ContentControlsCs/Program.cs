using GemBox.Document;
using GemBox.Document.CustomMarkups;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

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
}