Imports System
Imports GemBox.Document
Imports GemBox.Document.Tables

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As DocumentModel = New DocumentModel

        ' MS word uses a sequence of 5 EN SPACE characters for empty FORMTEXT field and so will this sample.
        Dim formTextPlaceholder = New String(ChrW(&H2002), 5)

        ' Table with 2 columns will be used to layout labels and form fields.
        ' First column will contain label and second column will contain form field.
        document.Sections.Add(
            New Section(document,
                New Paragraph(document, "Form"),
                    New Table(document,
                        New TableRow(document,
                            New TableCell(document,
                                New Paragraph(document, "Full name:")),
                            New TableCell(document,
                                New Paragraph(document,
                                    New Field(document, FieldType.FormText, Nothing, formTextPlaceholder)))),
                        New TableRow(document,
                            New TableCell(document,
                                New Paragraph(document, "Birth date:")),
                            New TableCell(document,
                                New Paragraph(document,
                                    New Field(document, FieldType.FormText, Nothing, formTextPlaceholder)))),
                        New TableRow(document,
                            New TableCell(document,
                                New Paragraph(document, "Salary:")),
                            New TableCell(document,
                                New Paragraph(document,
                                    New Run(document, "$"),
                                    New Field(document, FieldType.FormText, Nothing, formTextPlaceholder)))),
                        New TableRow(document,
                            New TableCell(document,
                                New Paragraph(document, "Married:")),
                            New TableCell(document,
                                New Paragraph(document,
                                    New Field(document, FieldType.FormCheckBox)))),
                        New TableRow(document,
                            New TableCell(document,
                                New Paragraph(document, "Gender:")),
                            New TableCell(document,
                                New Paragraph(document,
                                    New Field(document, FieldType.FormDropDown)))))))

        ' Set form field names.
        Dim fieldNames = New String() {"FullName", "BirthDate", "Salary", "Married", "Gender"}
        Dim i As Integer = 0
        For Each field As Field In document.GetChildElements(True, ElementType.Field)
            field.FormData.Name = fieldNames(i)
            i += 1
        Next

        ' Format heading paragraph.
        Dim heading = document.Sections(0).Blocks.Cast(Of Paragraph)(0)
        heading.ParagraphFormat.Style = DirectCast(document.Styles.GetOrAdd(StyleTemplateType.Heading1), ParagraphStyle)
        heading.ParagraphFormat.Alignment = HorizontalAlignment.Center

        ' Format table.
        Dim table = document.Sections(0).Blocks.Cast(Of Table)(1)
        table.TableFormat.PreferredWidth = New TableWidth(80, TableWidthUnit.Percentage)
        table.TableFormat.Alignment = HorizontalAlignment.Center
        table.TableFormat.Borders.SetBorders(MultipleBorderTypes.Inside, BorderStyle.None, Color.Empty, 0)
        table.TableFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Double, Color.Black, 1)

        ' Format table rows and cells.
        For Each row In table.Rows
            row.RowFormat.Height = New TableRowHeight(
        LengthUnitConverter.Convert(2, LengthUnit.Centimeter, LengthUnit.Point),
        TableRowHeightRule.Exact)

            Dim labelCell = row.Cells(0)
            labelCell.CellFormat.PreferredWidth = New TableWidth(50, TableWidthUnit.Percentage)
            labelCell.CellFormat.VerticalAlignment = VerticalAlignment.Center
            labelCell.Blocks.Cast(Of Paragraph)(0).ParagraphFormat.Alignment = HorizontalAlignment.Right

            Dim fieldCell = row.Cells(1)
            fieldCell.CellFormat.PreferredWidth = New TableWidth(50, TableWidthUnit.Percentage)
            fieldCell.CellFormat.VerticalAlignment = VerticalAlignment.Center
            fieldCell.Blocks.Cast(Of Paragraph)(0).ParagraphFormat.Alignment = HorizontalAlignment.Left
        Next

        ' Customize form fields.
        Dim formFieldsData = table.Content.FormFieldsData

        Dim fullNameFieldData = DirectCast(formFieldsData("FullName"), FormTextData)
        fullNameFieldData.MaximumLength = 50
        fullNameFieldData.ValueFormat = "Title case"
        ' Status text shows in lower right corner of MS Word (in status bar) and 
        ' help text shows if focus is on form field and F1 is pressed.
        fullNameFieldData.HelpText = "Enter your name and surname (trimmed to 50 characters)."
        fullNameFieldData.StatusText = "Enter your name and surname (trimmed to 50 characters)."

        Dim birthdateFieldData = DirectCast(formFieldsData("BirthDate"), FormTextData)
        birthdateFieldData.TextType = FormTextType.[Date]
        birthdateFieldData.ValueFormat = "yyyy-MM-dd"
        birthdateFieldData.HelpText = "Enter your date of birth."
        birthdateFieldData.StatusText = "Enter your date of birth."

        Dim salaryFieldData = DirectCast(formFieldsData("Salary"), FormTextData)
        salaryFieldData.TextType = FormTextType.Number
        salaryFieldData.ValueFormat = "#,##0.00"
        salaryFieldData.HelpText = "Enter your monthly salary in USD."
        salaryFieldData.StatusText = "Enter your monthly salary in USD."

        Dim marriedFieldData = DirectCast(formFieldsData("Married"), FormCheckBoxData)
        marriedFieldData.HelpText = "Mark as checked if you are married."
        marriedFieldData.StatusText = "Mark as checked if you are married."

        Dim genderFieldData = DirectCast(formFieldsData("Gender"), FormDropDownData)
        ' This is default option which signifies that user hasn't selected any gender.
        genderFieldData.Items.Add("<Select gender>")
        genderFieldData.Items.Add("Male")
        genderFieldData.Items.Add("Female")
        genderFieldData.HelpText = "Select your gender."
        genderFieldData.StatusText = "Select your gender."

        ' Make document a form - restrict editing to filling-in form fields.
        document.Protection.StartEnforcingProtection(EditingRestrictionType.FillingForms, "pass")

        document.Save("Create Form.docx")

    End Sub

End Module