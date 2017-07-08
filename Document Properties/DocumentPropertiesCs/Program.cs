using System;
using GemBox.Document;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        DocumentModel document = DocumentModel.Load("Reading.docx");

        Paragraph builtInPara = new Paragraph(document,
            new Run(document, "Built-in document properties:"),
            new SpecialCharacter(document, SpecialCharacterType.LineBreak));
        builtInPara.ParagraphFormat.Alignment = HorizontalAlignment.Left;

        foreach (var docProp in document.DocumentProperties.BuiltIn)
        {
            builtInPara.Inlines.Add(
                new Run(document, string.Format("{0}: {1}", docProp.Key, docProp.Value)));

            builtInPara.Inlines.Add(new SpecialCharacter(document, SpecialCharacterType.LineBreak));
        }

        Paragraph customPropPara = new Paragraph(document,
            new Run(document, "Custom document properties:"),
            new SpecialCharacter(document, SpecialCharacterType.LineBreak));
        customPropPara.ParagraphFormat.Alignment = HorizontalAlignment.Left;

        foreach (var docProp in document.DocumentProperties.Custom)
        {
            customPropPara.Inlines.Add(
                new Run(document, string.Format("{0}: {1} (Type: {2})", docProp.Key, docProp.Value, docProp.Value.GetType())));

            customPropPara.Inlines.Add(new SpecialCharacter(document, SpecialCharacterType.LineBreak));
        }

        document.Sections.Clear();
        document.Sections.Add(new Section(document, builtInPara, customPropPara));

        document.Save("Document Properties.docx");
    }
}
