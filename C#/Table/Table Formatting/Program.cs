using GemBox.Document;
using GemBox.Document.Tables;

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

        var table = new Table(document);
        table.TableFormat.AutomaticallyResizeToFitContents = false;

        // By default Table has assigned "Table Grid" style, the same as when creating it in Microsoft Word.
        // This base style defines borders which can be removed with the following.
        table.TableFormat.Style.TableFormat.Borders.ClearBorders();

        // Add columns with specified width.
        table.Columns.Add(new TableColumn(60));
        table.Columns.Add(new TableColumn(120));
        table.Columns.Add(new TableColumn(180));

        // Add rows with specified height.
        table.Rows.Add(new TableRow(document) { RowFormat = { Height = new TableRowHeight(30, TableRowHeightRule.AtLeast) } });
        table.Rows.Add(new TableRow(document) { RowFormat = { Height = new TableRowHeight(60, TableRowHeightRule.AtLeast) } });
        table.Rows.Add(new TableRow(document) { RowFormat = { Height = new TableRowHeight(90, TableRowHeightRule.AtLeast) } });

        for (int r = 0; r < 3; r++)
            for (int c = 0; c < 3; c++)
            {
                // Add cell.
                var cell = new TableCell(document);
                table.Rows[r].Cells.Add(cell);

                // Set cell's vertical alignment.
                switch (r)
                {
                    case 0:
                        cell.CellFormat.VerticalAlignment = VerticalAlignment.Top;
                        break;
                    case 1:
                        cell.CellFormat.VerticalAlignment = VerticalAlignment.Center;
                        break;
                    case 2:
                        cell.CellFormat.VerticalAlignment = VerticalAlignment.Bottom;
                        break;
                }

                // Add cell content.
                var paragraph = new Paragraph(document, $"Cell ({r + 1},{c + 1})");
                cell.Blocks.Add(paragraph);

                // Set cell content's horizontal alignment.
                switch (c)
                {
                    case 0:
                        paragraph.ParagraphFormat.Alignment = HorizontalAlignment.Left;
                        break;
                    case 1:
                        paragraph.ParagraphFormat.Alignment = HorizontalAlignment.Center;
                        break;
                    case 2:
                        paragraph.ParagraphFormat.Alignment = HorizontalAlignment.Right;
                        break;
                }

                if ((r + c) % 2 == 0)
                {
                    // Set cell's background and borders.
                    cell.CellFormat.BackgroundColor = new Color(255, 242, 204);
                    cell.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Double, Color.Red, 1);
                }
            }

        document.Sections.Add(new Section(document, table));

        document.Save("Table Formatting.docx");
    }

    static void Example2()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = new DocumentModel();

        var section = new Section(document);
        document.Sections.Add(section);

        section.Blocks.Add(new Paragraph(document,
            new Run(document, "Keep table on same page.") { CharacterFormat = { Size = 36 } },
            new SpecialCharacter(document, SpecialCharacterType.LineBreak),
            new Run(document, "This paragraph has a large spacing before to occupy most of the page.") { CharacterFormat = { Size = 14 } })
        { ParagraphFormat = { SpaceBefore = 400 } });

        var table = new Table(document, 20, 4);
        table.TableFormat.PreferredWidth = new TableWidth(100, TableWidthUnit.Percentage);
        section.Blocks.Add(table);

        // If you were to save a document at this point, you'd notice that the last few rows don't fit on the same page.
        // In other words, the table rows break across the first and second page.
        //document.Save("TableOnTwoPages.docx");

        // To prevent the table breaking across two pages, you need to set the KeepWithNext formatting.
        foreach (TableCell cell in table.GetChildElements(true, ElementType.TableCell))
        {
            // Cell should have at least one paragraph.
            if (cell.Blocks.Count == 0)
                cell.Blocks.Add(new Paragraph(cell.Document));

            foreach (Paragraph paragraph in cell.GetChildElements(true, ElementType.Paragraph))
                paragraph.ParagraphFormat.KeepWithNext = true;
        }

        document.Save("TableOnOnePage.docx");
    }
}
