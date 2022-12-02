using GemBox.Document;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        Example1();
        Example2();
    }

    static void Example1()
    {
        var document = DocumentModel.Load("MergeCustomizations.docx");

        document.MailMerge.FieldMerging += (sender, e) =>
        {
            if (e.IsValueFound && e.Value != null)
            {
                switch (e.FieldName)
                {
                    case "CheckedField":
                        bool checkedValue = (bool)e.Value;
                        var run = (Run)e.Inline;
                        run.CharacterFormat.FontColor = checkedValue ? Color.Green : Color.Red;
                        run.Text = checkedValue ? "☑" : "☒";
                        break;

                    case "LinkField":
                        var linkValue = ((string Address, string DisplayText))e.Value;
                        e.Inline = new Hyperlink(e.Document, linkValue.Address, linkValue.DisplayText);
                        break;

                    case "ImageField":
                        var imagePath = e.Value.ToString();
                        e.Inline = new Picture(e.Document, imagePath);
                        break;
                }
            }
        };

        document.MailMerge.Execute(
            new
            {
                CheckedField = true,
                LinkField = (Address: "https://www.gemboxsoftware.com/", DisplayText: "GemBox Homepage"),
                ImageField = "Dices.png"
            });

        document.Save("Merged Customizations Output.docx");
    }

    static void Example2()
    {
        var document = DocumentModel.Load("MergeCustomTypes.docx");

        // Map merge fields with prefix to data source without prefix.
        // E.g. "Html:MyName" field's name to "MyName" data source name.
        foreach (string fieldName in document.MailMerge.GetMergeFieldNames())
        {
            int index = fieldName.IndexOf(':');
            if (index > 0 && !document.MailMerge.FieldMappings.ContainsKey(fieldName))
                document.MailMerge.FieldMappings.Add(fieldName, fieldName.Substring(index + 1));
        }

        // Customize mail merge to support our custom prefixes.
        document.MailMerge.FieldMerging += (sender, e) =>
        {
            if (!e.IsValueFound)
                return;

            bool customImport = true;

            if (e.FieldName.StartsWith("Html:"))
                e.Field.Content.End.LoadText((string)e.Value, LoadOptions.HtmlDefault);
            else if (e.FieldName.StartsWith("Rtf:"))
                e.Field.Content.End.LoadText((string)e.Value, LoadOptions.RtfDefault);
            else if (e.FieldName.StartsWith("Docx:"))
                e.Field.Content.End.InsertRange(DocumentModel.Load((string)e.Value).Sections[0].Blocks.Content);
            else
                customImport = false;

            if (customImport)
            {
                // Remove the default import.
                e.Inline = null;

                // Check if the content is inserted into the same parent paragraph or added after it.
                // This depends on whether the data source contains inline-level or block-level elements.
                // If it's added after, then remove the empty parent which used to contain merge field.
                if (e.Field.ParentCollection.Count == 1)
                    e.Field.Parent.Content.Delete();
            }
        };

        document.MailMerge.Execute(
            new
            {
                Field1 = @"{\rtf1\ansi\deff0{\fonttbl{\f0 Arial Black;}}{\colortbl ;\red255\green128\blue64;}\f0\cf1 This is rich formatted text (in RTF format).}",
                Field2 = "<p style='font-family:Arial Narrow;color:royalblue;'>This is another rich formatted text (in HTML format).</p>",
                Field3 = "<p style='font-family:Arial Narrow;color:seagreen;'>And another rich formatted text (in HTML format).</p>",
                Field4 = "Reading.docx"
            });

        document.Save("Merged Custom Types Output.docx");
    }
}