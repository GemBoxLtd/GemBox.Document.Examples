Imports System.Linq
Imports GemBox.Document

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Example1()
        Example2()

    End Sub

    Sub Example1()
        Dim document As New DocumentModel()

        Dim section As New Section(document)
        document.Sections.Add(section)

        Dim blocks = section.Blocks

        ' Create bullet list style.
        Dim bulletList As New ListStyle(ListTemplateType.Bullet)
        bulletList.ListLevelFormats(0).ParagraphFormat.NoSpaceBetweenParagraphsOfSameStyle = True
        bulletList.ListLevelFormats(0).CharacterFormat.FontColor = Color.Red

        ' Create bullet list items.
        blocks.Add(New Paragraph(document, "First item.") With
        {
            .ListFormat = New ListFormat() With {.Style = bulletList}
        })
        blocks.Add(New Paragraph(document, "Second item.") With
        {
            .ListFormat = New ListFormat() With {.Style = bulletList}
        })
        blocks.Add(New Paragraph(document, "Third item.") With
        {
            .ListFormat = New ListFormat() With {.Style = bulletList}
        })

        blocks.Add(New Paragraph(document))

        ' Create number list style.
        Dim numberList As New ListStyle(ListTemplateType.NumberWithDot)

        ' Create number list items.
        blocks.Add(New Paragraph(document, "First item.") With
        {
            .ListFormat = New ListFormat() With {.Style = numberList}
        })
        blocks.Add(New Paragraph(document, "Sub item 1. a.") With
        {
            .ListFormat = New ListFormat() With {.Style = numberList, .ListLevelNumber = 1}
        })
        blocks.Add(New Paragraph(document, "Item below sub item 1. a.") With
        {
            .ListFormat = New ListFormat() With {.Style = numberList, .ListLevelNumber = 2}
        })
        blocks.Add(New Paragraph(document, "Sub item 1. b.") With
        {
            .ListFormat = New ListFormat() With {.Style = numberList, .ListLevelNumber = 1}
        })
        blocks.Add(New Paragraph(document, "Second item.") With
        {
            .ListFormat = New ListFormat() With {.Style = numberList}
        })

        document.Save("Lists.docx")
    End Sub

    Sub Example2()
        Dim listItemsCount As Integer = 12
        
        Dim document As New DocumentModel()

        Dim section As New Section(document)
        document.Sections.Add(section)

        ' Create number list style.
        Dim numberList As New ListStyle(ListTemplateType.NumberWithDot)

        ' Customize list level formats.
        For level = 0 To numberList.ListLevelFormats.Count - 1

            Dim levelFormat As ListLevelFormat = numberList.ListLevelFormats(level)

            levelFormat.ParagraphFormat.NoSpaceBetweenParagraphsOfSameStyle = True
            levelFormat.Alignment = HorizontalAlignment.Left
            levelFormat.NumberStyle = NumberStyle.Decimal

            levelFormat.NumberPosition = 18 * level
            levelFormat.NumberFormat = String.Concat(Enumerable.Range(1, level + 1).Select(Function(i) $"%{i}."))

        Next

        ' Create number list items.
        For i = 0 To listItemsCount - 1

            Dim paragraph As New Paragraph(document, "Lorem ipsum")

            paragraph.ListFormat.Style = numberList
            paragraph.ListFormat.ListLevelNumber = i Mod 9

            section.Blocks.Add(paragraph)

        Next

        document.Save("CustomizedList.docx")
    End Sub

End Module