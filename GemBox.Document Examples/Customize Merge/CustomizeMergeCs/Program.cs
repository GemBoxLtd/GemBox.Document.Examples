using GemBox.Document;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

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
}