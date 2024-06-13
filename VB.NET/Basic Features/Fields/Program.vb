Imports GemBox.Document

Module Program

    Sub Main()
        Example1()
        Example2()
    End Sub

    Sub Example1()
        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As New DocumentModel()

        Dim section As New Section(document)
        document.Sections.Add(section)

        ' Add '{ AUTHOR }' field.
        section.Blocks.Add(
            New Paragraph(document,
                New Run(document, "Author: "),
                New Field(document, FieldType.Author, Nothing, "Mario at GemBox")))

        ' Add '{ DATE }' field.
        section.Blocks.Add(
            New Paragraph(document,
                New Run(document, "Date: "),
                New Field(document, FieldType.Date)))

        ' Add '{ DATE \@ "dddd, MMMM dd, yyyy" }' field.
        section.Blocks.Add(
            New Paragraph(document,
                New Run(document, "Date with specified format: "),
                New Field(document, FieldType.Date, "\@ ""dddd, MMMM dd, yyyy""")))

        document.Save("Fields.docx")
    End Sub

    Sub Example2()
        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document = New DocumentModel()

        Dim section = New Section(document)
        document.Sections.Add(section)

        ' A simple '{ IF }' field, that will result on different texts if true or false.
        Dim simpleField = New Field(document, FieldType.If, "10 = 10 ""The result Is True."" ""The result Is False.""")

        ' A complex '{ IF }' field, that will result on texts with different formats if true or false.
        Dim complexField = New Field(document, FieldType.If,
            New Run(document, "10 = 10 "),
            New Run(document, """The result Is True.""") With {.CharacterFormat = New CharacterFormat() With {.FontColor = Color.Green, .Bold = True}},
            New Run(document, """The result Is False.""") With {.CharacterFormat = New CharacterFormat() With {.FontColor = Color.Red, .Italic = True}})

        ' Add both fields to the document.
        section.Blocks.Add(New Paragraph(document, simpleField))
        section.Blocks.Add(New Paragraph(document, complexField))

        ' Call Update method to resolve the field's value.
        simpleField.Update()
        complexField.Update()

        ' Get the result, which for simpleField will be a simple Run with "The result is true." as text.
        Dim simpleResult = TryCast(simpleField.ResultInlines.First(), Run)

        ' Get the result, which will be a Run with FontColor Green And Bold set to true.
        Dim complexResult = TryCast(complexField.ResultInlines.First(), Run)

        document.Save("FieldsUpdated.docx")
    End Sub

End Module
