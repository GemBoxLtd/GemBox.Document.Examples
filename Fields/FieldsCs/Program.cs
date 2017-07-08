using System;
using GemBox.Document;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        DocumentModel document = new DocumentModel();

        document.Sections.Add(
            new Section(document,
                new Paragraph(document, "Press Alt + F9 to see field codes!"),
                new Paragraph(document,
                    new Run(document, "Date: "),
                    // { DATE }
                    new Field(document, FieldType.Date),
                    new SpecialCharacter(document, SpecialCharacterType.LineBreak),
                    new Run(document, "Date (formatted): "),
                    // { DATE \@ "dddd, MMMM dd, yyyy"  \* MERGEFORMAT }
                    new Field(document, FieldType.Date, "\\@ \"dddd, MMMM dd, yyyy\"  \\* MERGEFORMAT"),
                    new SpecialCharacter(document, SpecialCharacterType.LineBreak),
                    new Run(document, "Date & Time (formatted): "),
                    // { DATE  \@ "M/d/yyyy h:mm:ss am/pm"  \* MERGEFORMAT }
                    new Field(document, FieldType.Date, " \\@ \"M/d/yyyy h:mm:ss am/pm\"  \\* MERGEFORMAT"))));

        document.Save("Fields.docx");
    }
}
