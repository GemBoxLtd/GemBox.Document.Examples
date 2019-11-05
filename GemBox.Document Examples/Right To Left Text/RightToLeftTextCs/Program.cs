using GemBox.Document;
using GemBox.Document.Tables;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = DocumentModel.Load("RightToLeft.docx");

        // Show line numbers on the right side of the page
        var pageSetup = document.Sections[0].PageSetup;
        pageSetup.LineNumberRestartSetting = LineNumberRestartSetting.Continuous;
        pageSetup.RightToLeft = true;

        // Create a new right-to-left paragraph
        var paragraph = new Paragraph(document);
        paragraph.ParagraphFormat.RightToLeft = true;
        paragraph.Inlines.Add(new Run(document, "أخذ عن موالية الإمتعاض"));
        document.Sections[0].Blocks.Add(paragraph);

        // Create a right-to-left table
        var table = new Table(document);
        table.TableFormat.RightToLeft = true;
        table.TableFormat.PreferredWidth = new TableWidth(100, TableWidthUnit.Percentage);
        var row = new TableRow(document);
        table.Rows.Add(row);

        var firstCellPara = new Paragraph(document, "של תיבת תרומה מלא");
        firstCellPara.ParagraphFormat.RightToLeft = true;
        row.Cells.Add(new TableCell(document, firstCellPara));

        var secondCellPara = new Paragraph(document, "200");
        row.Cells.Add(new TableCell(document, secondCellPara));

        document.Sections[0].Blocks.Add(table);

        document.Save("RightToLeft.pdf");
    }
}