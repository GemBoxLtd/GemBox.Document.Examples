using GemBox.Document;

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

        // Word files that will be combined into one file.
        string[] files =
        {
            "MergeFile01.docx",
            "MergeFile02.docx",
            "MergeFile03.docx"
        };

        // Create destination document.
        var destination = new DocumentModel();

        // Merge multiple source documents by importing their content at the end.
        foreach (var file in files)
        {
            var source = DocumentModel.Load(file);
            destination.Content.End.InsertRange(source.Content);
        }

        // Save joined documents into one file.
        destination.Save("Merged Files.docx");
    }

    static void Example2()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        // Word files that will be combined into one file.
        string[] files =
        {
            "MergeFile01.docx",
            "MergeFile02.docx",
            "MergeFile03.docx"
        };

        var destination = new DocumentModel();
        var firstSourceDocument = true;

        foreach (var file in files)
        {
            var source = DocumentModel.Load(file);
            var firstSourceSection = true;

            // Reuse the same mapping for importing to improve performance.
            var mapping = new ImportMapping(source, destination, false);

            foreach (var sourceSection in source.Sections)
            {
                // Import section from source document to destination document.
                var destinationSection = destination.Import(sourceSection, true, mapping);
                destination.Sections.Add(destinationSection);

                // Set the first section to start on the same page as the previous section.
                // In other words, the source content continues to flow with the current destination content.
                if (firstSourceSection)
                {
                    destinationSection.PageSetup.SectionStart = SectionStart.Continuous;
                    firstSourceSection = false;
                }
            }

            // Set the destination's default formatting to first source's default formatting.
            // Note, a single document can only have one default formatting.
            if (firstSourceDocument)
            {
                destination.DefaultCharacterFormat = source.DefaultCharacterFormat.Clone();
                destination.DefaultParagraphFormat = source.DefaultParagraphFormat.Clone();
                firstSourceDocument = false;
            }
        }

        // Save joined sections into one file.
        destination.Save("Merged Sections.docx");
    }
}
