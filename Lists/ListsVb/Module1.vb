Imports System
Imports GemBox.Document

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As DocumentModel = New DocumentModel

        ' First create list style.
        Dim numberList As New ListStyle(ListTemplateType.NumberWithDot)

        ' Then set Paragraph.ListFormat.Style and ParagraphListFormat.ListLevelNumber.
        Dim para1 As New Paragraph(document, "First item.")
        para1.ListFormat.Style = numberList
        para1.ParagraphFormat.NoSpaceBetweenParagraphsOfSameStyle = True

        Dim para2 As New Paragraph(document, "Second item.")
        para2.ListFormat.Style = numberList
        para2.ParagraphFormat.NoSpaceBetweenParagraphsOfSameStyle = True

        Dim para2a As New Paragraph(document, "Sub item 2a.")
        para2a.ListFormat.Style = numberList
        para2a.ListFormat.ListLevelNumber += 1
        para2a.ParagraphFormat.NoSpaceBetweenParagraphsOfSameStyle = True

        Dim para2b As New Paragraph(document, "Sub item 2b.")
        para2b.ListFormat.Style = numberList
        para2b.ListFormat.ListLevelNumber += 1
        para2b.ParagraphFormat.NoSpaceBetweenParagraphsOfSameStyle = True

        Dim para3 As New Paragraph(document, "Third item.")
        para3.ListFormat.Style = numberList
        para3.ParagraphFormat.NoSpaceBetweenParagraphsOfSameStyle = True

        Dim section As New Section(document, para1, para2, para2a, para2b, para3)
        document.Sections.Add(section)

        section.Blocks.Add(New Paragraph(document))

        ' Create list style.
        Dim bulletList As New ListStyle(ListTemplateType.Bullet)

        ' Use it.
        section.Blocks.Add(
            New Paragraph(document, "First item.") With {
                .ParagraphFormat = New ParagraphFormat() With {
                    .NoSpaceBetweenParagraphsOfSameStyle = True
                },
                .ListFormat = New ListFormat() With {
                    .Style = bulletList
                }})

        section.Blocks.Add(
            New Paragraph(document, "Second item.") With {
                .ParagraphFormat = New ParagraphFormat() With {
                    .NoSpaceBetweenParagraphsOfSameStyle = True
                },
                .ListFormat = New ListFormat() With {
                    .Style = bulletList
                }})

        section.Blocks.Add(
            New Paragraph(document, "Sub item 2a.") With {
                .ParagraphFormat = New ParagraphFormat() With {
                    .NoSpaceBetweenParagraphsOfSameStyle = True
                },
                .ListFormat = New ListFormat() With {
                    .Style = bulletList,
                    .ListLevelNumber = 1
                }})

        section.Blocks.Add(
            New Paragraph(document, "Item below sub item 2a.") With {
                .ParagraphFormat = New ParagraphFormat() With {
                    .NoSpaceBetweenParagraphsOfSameStyle = True
                },
                .ListFormat = New ListFormat() With {
                    .Style = bulletList,
                    .ListLevelNumber = 2
                }})

        section.Blocks.Add(
            New Paragraph(document, "Third item.") With {
                .ParagraphFormat = New ParagraphFormat() With {
                    .NoSpaceBetweenParagraphsOfSameStyle = True
                },
                .ListFormat = New ListFormat() With {
                    .Style = bulletList
                }})

        document.Save("Lists.docx")

    End Sub

End Module