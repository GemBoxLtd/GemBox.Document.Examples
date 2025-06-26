using GemBox.Document;
using System.Linq;

class Program
{
    static void Main()
    {
        Example1();
        Example2();
    }

    static void Example1()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = new DocumentModel();

        var section = new Section(document);
        document.Sections.Add(section);

        // Add '{ AUTHOR }' field.
        section.Blocks.Add(
            new Paragraph(document,
                new Run(document, "Author: "),
                new Field(document, FieldType.Author, null, "Mario at GemBox")));

        // Add '{ DATE }' field.
        section.Blocks.Add(
            new Paragraph(document,
                new Run(document, "Date: "),
                new Field(document, FieldType.Date)));

        // Add '{ DATE \@ "dddd, MMMM dd, yyyy" }' field.
        section.Blocks.Add(
            new Paragraph(document,
                new Run(document, "Date with specified format: "),
                new Field(document, FieldType.Date, "\\@ \"dddd, MMMM dd, yyyy\"")));

        document.Save("Fields.docx");
    }

    static void Example2()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = new DocumentModel();

        var section = new Section(document);
        document.Sections.Add(section);

        // A simple '{ IF }' field, that will result on different texts if true or false.
        var simpleField = new Field(document, FieldType.If, "10 = 10 \"The result is true.\" \"The result is false.\"");

        // A complex '{ IF }' field, that will result on texts with different formats if true or false.
        var complexField = new Field(document, FieldType.If,
            new Run(document, "10 = 10 "),
            new Run(document, "\"The result is true.\"") { CharacterFormat = { FontColor = Color.Green, Bold = true } },
            new Run(document, "\"The result is false.\"") { CharacterFormat = { FontColor = Color.Red, Italic = true } });

        // Add both fields to the document.
        section.Blocks.Add(new Paragraph(document, simpleField));
        section.Blocks.Add(new Paragraph(document, complexField));

        // Call Update method to resolve the field's value.
        simpleField.Update();
        complexField.Update();

        // Get the result, which for simpleField will be a simple Run with "The result is true." as text.
        Run simpleResult = simpleField.ResultElements.OfType<Run>().First();

        // Get the result, which will be a Run with FontColor Green and Bold set to true.
        Run complexResult = complexField.ResultElements.OfType<Run>().First();

        document.Save("FieldsUpdated.docx");
    }
}
