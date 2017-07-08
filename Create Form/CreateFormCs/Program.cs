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

        // MS word uses a sequence of 5 EN SPACE characters for empty FORMTEXT field and so will this sample.
        var formTextPlaceholder = new string('\x2002', 5);

        // Table with 2 columns will be used to layout labels and form fields.
        // First column will contain label and second column will contain form field.
        document.Sections.Add(
            new Section(document,
                new Paragraph(document, "Form"),
                new Table(document,
                    new TableRow(document,
                        new TableCell(document,
                            new Paragraph(document, "Full name:")),
                        new TableCell(document,
                            new Paragraph(document,
                                new Field(document, FieldType.FormText, null, formTextPlaceholder)
                                { FormData = { Name = "FullName" } }))),
                    new TableRow(document,
                        new TableCell(document,
                            new Paragraph(document, "Birth date:")),
                        new TableCell(document,
                            new Paragraph(document,
                                new Field(document, FieldType.FormText, null, formTextPlaceholder)
                                { FormData = { Name = "BirthDate" } }))),
                    new TableRow(document,
                        new TableCell(document,
                            new Paragraph(document, "Salary:")),
                        new TableCell(document,
                            new Paragraph(document,
                                new Run(document, "$"),
                                new Field(document, FieldType.FormText, null, formTextPlaceholder)
                                { FormData = { Name = "Salary" } }))),
                    new TableRow(document,
                        new TableCell(document,
                            new Paragraph(document, "Married:")),
                        new TableCell(document,
                            new Paragraph(document,
                                new Field(document, FieldType.FormCheckBox)
                                { FormData = { Name = "Married" } }))),
                    new TableRow(document,
                        new TableCell(document,
                            new Paragraph(document, "Gender:")),
                        new TableCell(document,
                            new Paragraph(document,
                                new Field(document, FieldType.FormDropDown)
                                { FormData = { Name = "Gender" } }))))));

        // Format heading paragraph.
        var heading = document.Sections[0].Blocks.Cast<Paragraph>(0);
        heading.ParagraphFormat.Style = (ParagraphStyle)document.Styles.GetOrAdd(StyleTemplateType.Heading1);
        heading.ParagraphFormat.Alignment = HorizontalAlignment.Center;

        // Format table.
        var table = document.Sections[0].Blocks.Cast<Table>(1);
        table.TableFormat.PreferredWidth = new TableWidth(80, TableWidthUnit.Percentage);
        table.TableFormat.Alignment = HorizontalAlignment.Center;
        table.TableFormat.Borders.SetBorders(MultipleBorderTypes.Inside, BorderStyle.None, Color.Empty, 0);
        table.TableFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Double, Color.Black, 1);

        // Format table rows and cells.
        foreach (var row in table.Rows)
        {
            row.RowFormat.Height = new TableRowHeight(
                LengthUnitConverter.Convert(2, LengthUnit.Centimeter, LengthUnit.Point),
                TableRowHeightRule.Exact);

            var labelCell = row.Cells[0];
            labelCell.CellFormat.PreferredWidth = new TableWidth(50, TableWidthUnit.Percentage);
            labelCell.CellFormat.VerticalAlignment = VerticalAlignment.Center;
            labelCell.Blocks.Cast<Paragraph>(0).ParagraphFormat.Alignment = HorizontalAlignment.Right;

            var fieldCell = row.Cells[1];
            fieldCell.CellFormat.PreferredWidth = new TableWidth(50, TableWidthUnit.Percentage);
            fieldCell.CellFormat.VerticalAlignment = VerticalAlignment.Center;
            fieldCell.Blocks.Cast<Paragraph>(0).ParagraphFormat.Alignment = HorizontalAlignment.Left;
        }

        // Customize form fields.
        var formFieldsData = table.Content.FormFieldsData;

        var fullNameFieldData = (FormTextData)formFieldsData["FullName"];
        fullNameFieldData.MaximumLength = 50;
        fullNameFieldData.ValueFormat = "Title case";
        // Status text shows in lower right corner of MS Word (in status bar) and 
        // help text shows if focus is on form field and F1 is pressed.
        fullNameFieldData.StatusText = fullNameFieldData.HelpText =
            "Enter your name and surname (trimmed to 50 characters).";

        var birthdateFieldData = (FormTextData)formFieldsData["BirthDate"];
        birthdateFieldData.TextType = FormTextType.Date;
        birthdateFieldData.ValueFormat = "yyyy-MM-dd";
        birthdateFieldData.StatusText = birthdateFieldData.HelpText =
            "Enter your date of birth.";

        var salaryFieldData = (FormTextData)formFieldsData["Salary"];
        salaryFieldData.TextType = FormTextType.Number;
        salaryFieldData.ValueFormat = "#,##0.00";
        salaryFieldData.StatusText = salaryFieldData.HelpText =
            "Enter your monthly salary in USD.";

        var marriedFieldData = (FormCheckBoxData)formFieldsData["Married"];
        marriedFieldData.StatusText = marriedFieldData.HelpText =
            "Mark as checked if you are married.";

        var genderFieldData = (FormDropDownData)formFieldsData["Gender"];
        // This is default option which signifies that user hasn't selected any gender.
        genderFieldData.Items.Add("<Select gender>");
        genderFieldData.Items.Add("Male");
        genderFieldData.Items.Add("Female");
        genderFieldData.StatusText = genderFieldData.HelpText =
            "Select your gender.";

        // Make document a form - restrict editing to filling-in form fields.
        document.Protection.StartEnforcingProtection(EditingRestrictionType.FillingForms, "pass");

        document.Save("Create Form.docx");
    }
}
