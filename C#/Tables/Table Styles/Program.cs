using GemBox.Document;
using GemBox.Document.Tables;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = new DocumentModel();
        document.DefaultParagraphFormat.SpaceAfter = 0;

        var section = new Section(document);
        document.Sections.Add(section);

        section.Blocks.Add(new Paragraph(document, "Built-in table style (List Table 2 - Accent 2):"));

        // Create and add a built-in table style.
        var builtinTableStyle = (TableStyle)Style.CreateStyle(StyleTemplateType.ListTable2Accent2, document);
        document.Styles.Add(builtinTableStyle);

        // Modify built-in table style.
        builtinTableStyle.TableFormat.RowBandSize = 3;

        // Create a table and assign built-in table style to it.
        var tableWithBuiltinStyle = new Table(document, 10, 5,
            (int r, int c) => new TableCell(document, new Paragraph(document, $"Cell ({r + 1},{c + 1})")));
        section.Blocks.Add(tableWithBuiltinStyle);
        tableWithBuiltinStyle.TableFormat.PreferredWidth = new TableWidth(100, TableWidthUnit.Percentage);
        tableWithBuiltinStyle.TableFormat.Style = builtinTableStyle;
        tableWithBuiltinStyle.TableFormat.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.FirstColumn | TableStyleOptions.BandedRows;

        section.Blocks.Add(new Paragraph(document, "Custom table style (GemBox Table):") { ParagraphFormat = { SpaceBefore = 24 } });

        // Create and add a custom table style.
        var customTableStyle = new TableStyle("GemBox Table");
        document.Styles.Add(customTableStyle);

        // Set table style format.
        customTableStyle.TableFormat.Borders.SetBorders(MultipleBorderTypes.All, BorderStyle.Single, new Color(79, 129, 189), 1);
        customTableStyle.ParagraphFormat.Alignment = HorizontalAlignment.Center;

        // Set table style conditional format for first row.
        var firstRowFormat = customTableStyle.ConditionalFormats[TableStyleFormatType.FirstRow];
        firstRowFormat.CharacterFormat.Bold = true;
        firstRowFormat.CharacterFormat.FontColor = Color.White;
        firstRowFormat.ParagraphFormat.SpaceAfter = 12;
        firstRowFormat.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Single, new Color(79, 129, 189), 2.5);
        firstRowFormat.CellFormat.BackgroundColor = new Color(192, 80, 77);

        // Set table style conditional format for last row.
        var lastRowFormat = customTableStyle.ConditionalFormats[TableStyleFormatType.LastRow];
        lastRowFormat.CharacterFormat.Bold = true;
        lastRowFormat.ParagraphFormat.SpaceBefore = 12;
        lastRowFormat.CellFormat.Borders.SetBorders(MultipleBorderTypes.Top, BorderStyle.Double, new Color(79, 129, 189), 1.25);
        lastRowFormat.CellFormat.BackgroundColor = new Color(211, 223, 238);

        // Create a table and assign custom table style to it.
        var tableWithCustomStyle = new Table(document, 10, 5,
            (int r, int c) => new TableCell(document, new Paragraph(document, $"Cell ({r + 1},{c + 1})")));
        section.Blocks.Add(tableWithCustomStyle);
        tableWithCustomStyle.TableFormat.PreferredWidth = new TableWidth(100, TableWidthUnit.Percentage);
        tableWithCustomStyle.TableFormat.Style = customTableStyle;
        tableWithCustomStyle.TableFormat.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.LastRow;

        document.Save("Table Styles.docx");
    }
}
