Imports GemBox.Document

Module Program

    Sub Main()

        Example1()
        Example2()

    End Sub

    Sub Example1()
        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document = DocumentModel.Load("Input.docx")

        Dim paginator = document.GetPaginator()

        Dim secondPage = paginator.Pages(1)

        secondPage.Save("SecondPage.docx")
    End Sub

    Sub Example2()
        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document = DocumentModel.Load("Input.docx")

        Dim paginator = document.GetPaginator()

        For i = 0 To paginator.Pages.Count - 1
            Dim page = paginator.Pages(i)
            Dim pageRange = page.Range
            Dim start = pageRange.Start(0)

            Dim textBox = CreateTextBox(document, i)
            start.InsertRange(textBox.Content)
        Next

        document.Save("Output.docx")
    End Sub

    ' A floating textbox that will be inserted at the start of every page.
    Function CreateTextBox(document As DocumentModel, page As Integer) As TextBox
        Dim run = New Run(document, "Inserted textbox on page " & (page + 1))
        run.CharacterFormat.Size = 25
        run.CharacterFormat.FontColor = Color.White

        Dim textBox = New TextBox(document, New FloatingLayout(
            New HorizontalPosition(-340, LengthUnit.Point, HorizontalPositionAnchor.RightMargin),
            New VerticalPosition(0, LengthUnit.Point, VerticalPositionAnchor.Margin),
            New Size(340, 45, LengthUnit.Point)) With
        {.WrappingStyle = TextWrappingStyle.InFrontOfText})
        textBox.Fill.SetSolid(New Color(&H4472C4))
        textBox.Blocks.Add(New Paragraph(document, run))

        Return textBox
    End Function

End Module
