using System;
using GemBox.Document;
using GemBox.Document.Tables;
using GemBox.Document.Tracking;

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
        var document = DocumentModel.Load("Revisions.docx");
        var acceptRevisions = true;

        if (acceptRevisions)
            document.Revisions.AcceptAll();
        else
            document.Revisions.RejectAll();

        document.Save("Revised Document.docx");
    }

    static void Example2()
    {
        var document = DocumentModel.Load("Revisions.docx");

        // Iterate through all runs in the document.
        foreach (Run run in document.GetChildElements(true, ElementType.Run))
        {
            // Reject revision of the character format.
            if (run.CharacterFormatRevision != null)
                run.CharacterFormatRevision.Reject();

            // Reject deletion of a run.
            if (run.Revision?.RevisionType == RevisionType.Delete)
                run.Revision.Reject();
        }

        // Iterate through all remaining revisions in the document.
        foreach (var revision in document.Revisions)
        {
            // Accept only revisions from GemBox that were added last month.
            if (revision.Author == "GemBox" && revision.Date > DateTime.Now.AddMonths(-1))
                revision.Accept();
        }

        document.Save("Processed Revisions.docx");
    }

    static void Example3()
    {
        var document = DocumentModel.Load("NoRevisions.docx");
        var section = document.Sections[0];

        var paragraph1 = section.Blocks.Cast<Paragraph>(0);
        var run1 = paragraph1.Inlines.Cast<Run>(0);

        // 1. Changing the formatting of the run.
        var characterFormatRevision = new CharacterFormatRevision(document)
        {
            Author = "GemBox",
            Date = DateTime.Now
        };
        // CharacterFormatRevision.CharacterFormat holds the format which was used before the revision was applied.
        characterFormatRevision.CharacterFormat = run1.CharacterFormat.Clone();
        run1.CharacterFormatRevision = characterFormatRevision;
        // Changing the format.
        run1.CharacterFormat.UnderlineStyle = UnderlineType.Double;

        // 2. Removing the run.
        var run2 = paragraph1.Inlines.Cast<Run>(1);
        // Mark run as deleted.
        run2.Revision = new Revision(RevisionType.Delete) { Author = "GemBox" };

        // 3. Inserting a run.
        var run3 = new Run(document, "Run3");
        // Mark run as inserted.
        run3.Revision = new Revision(RevisionType.Insert) { Author = "GemBox" };
        paragraph1.Inlines.Add(run3);

        // 4. Joining paragraphs.
        var paragraph2 = section.Blocks.Cast<Paragraph>(1);
        // Marking paragraph as deleted doesn't remove the paragraph content, it joins it with the following paragraph.
        paragraph2.Revision = new Revision(RevisionType.Delete) { Author = "GemBox" };

        var table = section.Blocks.Cast<Table>(3);

        // 5. Removing a table row.
        var row2 = table.Rows[1];
        // Mark row as deleted.
        row2.Revision = new Revision(RevisionType.Delete) { Author = "GemBox" };

        // 6. Adding a new table row.
        var newRow = new TableRow(document,
            new TableCell(document, new Paragraph(document, "new row")),
            new TableCell(document));
        // Mark row as inserted.
        newRow.Revision = new Revision(RevisionType.Insert) { Author = "GemBox" };
        table.Rows.Add(newRow);

        document.Save("Added Revisions.docx");
    }
}