using System;
using GemBox.Document;
using GemBox.Document.Tables;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        DocumentModel document = new DocumentModel();

        int tableRowCount = 10;
        int tableColumnCount = 5;

        // Built-in table styles can be created using Style.CreateStyle() method.
        TableStyle lightShadingStyle = (TableStyle)Style.CreateStyle(StyleTemplateType.LightShadingAccent2, document);

        // We can also create our own (custom) table styles.
        TableStyle customTableStyle = new TableStyle("GemBox Table Style")
        {
            ParagraphFormat =
            {
                SpaceAfter = 0,
                LineSpacing = 1,
                LineSpacingRule = LineSpacingRule.Multiple
            },
            TableFormat =
            {
                RowBandSize = 2,
                ColumnBandSize = 1,
                IndentFromLeft = 0d,
                DefaultCellPadding = new Padding(108 / 20d, 0d)
            }
        };

        // Set custom style table borders.
        customTableStyle.TableFormat.Borders.SetBorders(MultipleBorderTypes.All, BorderStyle.Single, new Color(79, 129, 189), 1);

        // Table style conditional format - First row.
        var firstRowFormat = customTableStyle.ConditionalFormats[TableStyleFormatType.FirstRow];
        firstRowFormat.CharacterFormat.Bold = true;
        firstRowFormat.CharacterFormat.FontColor = Color.White;
        firstRowFormat.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Single, new Color(79, 129, 189), 1);
        firstRowFormat.CellFormat.Borders.SetBorders(MultipleBorderTypes.Bottom, BorderStyle.Single, new Color(79, 129, 189), 2.5);
        firstRowFormat.CellFormat.BackgroundColor = new Color(192, 80, 77);

        // Table style conditional format - Last row.
        var lastRowFormat = customTableStyle.ConditionalFormats[TableStyleFormatType.LastRow];
        lastRowFormat.CharacterFormat.Bold = true;
        lastRowFormat.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Single, new Color(79, 129, 189), 1);
        lastRowFormat.CellFormat.Borders.SetBorders(MultipleBorderTypes.Top, BorderStyle.Double, new Color(79, 129, 189), 1.25);

        // Table style conditional format - Banded rows.
        var bandedRowFormat = customTableStyle.ConditionalFormats[TableStyleFormatType.OddBandedRows];
        bandedRowFormat.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Single, new Color(79, 129, 189), 1);
        bandedRowFormat.CellFormat.BackgroundColor = new Color(211, 223, 238);

        // First add style to the document, then use it.
        document.Styles.Add(lightShadingStyle);
        document.Styles.Add(customTableStyle);

        document.Sections.Add(
            new Section(document,
                new Paragraph(document, "Built-in Table Style:"),
                new Table(document, tableRowCount, tableColumnCount,
                    (int rowIndex, int columnIndex) => new TableCell(document, new Paragraph(document, new Run(document, string.Format("{0}-{1}", rowIndex, columnIndex)))))
                {
                    TableFormat = new TableFormat()
                    {
                        PreferredWidth = new TableWidth(100, TableWidthUnit.Percentage),
                        Style = lightShadingStyle,
                        StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.FirstColumn | TableStyleOptions.BandedRows
                    }
                },
                new Paragraph(document),
                new Paragraph(document, "Custom Table Style:"),
                new Table(document, tableRowCount, tableColumnCount,
                    (int rowIndex, int columnIndex) => new TableCell(document, new Paragraph(document, new Run(document, string.Format("{0}-{1}", rowIndex, columnIndex)))))
                {
                    TableFormat = new TableFormat()
                    {
                        PreferredWidth = new TableWidth(100, TableWidthUnit.Percentage),
                        Style = customTableStyle,
                        StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.LastRow | TableStyleOptions.BandedRows
                    }
                }));

        document.Save("Table Styles.docx");
    }
}
