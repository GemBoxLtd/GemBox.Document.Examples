Imports System
Imports GemBox.Document
Imports GemBox.Document.Tables
Imports GemBox.Document.Tracking

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Example1()
        Example2()
        Example3()

    End Sub

    Sub Example1()
        Dim document = DocumentModel.Load("Revisions.docx")
        Dim acceptRevisions = True

        If acceptRevisions Then
            document.Revisions.AcceptAll()
        Else
            document.Revisions.RejectAll()
        End If

        document.Save("Revised Document.docx")
    End Sub

    Sub Example2()
        Dim document = DocumentModel.Load("Revisions.docx")

        ' Iterate through all runs in the document.
        For Each run As Run In document.GetChildElements(True, ElementType.Run)
            ' Reject revision of the character format.
            If run.CharacterFormatRevision IsNot Nothing Then
                run.CharacterFormatRevision.Reject()
            End If

            ' Reject deletion of a run.
            If run.Revision?.RevisionType = RevisionType.Delete Then
                run.Revision.Reject()
            End If
        Next

        ' Iterate through all remaining revisions in the document.
        For Each revision In document.Revisions
            ' Accept only revisions from GemBox that were added last month.
            If revision.Author = "GemBox" And revision.Date > DateTime.Now.AddMonths(-1) Then
                revision.Accept()
            End If
        Next

        document.Save("Processed Revisions.docx")
    End Sub

    Sub Example3()
        Dim document = DocumentModel.Load("NoRevisions.docx")
        Dim section = document.Sections(0)

        Dim paragraph1 = section.Blocks.Cast(Of Paragraph)(0)
        Dim run1 = paragraph1.Inlines.Cast(Of Run)(0)

        ' 1. Changing the formatting of the run.
        Dim characterFormatRevision = New CharacterFormatRevision(document) With
        {
            .Author = "GemBox",
            .Date = DateTime.Now
        }
        ' CharacterFormatRevision.CharacterFormat holds the format which was used before the revision was applied.
        characterFormatRevision.CharacterFormat = run1.CharacterFormat.Clone()
        run1.CharacterFormatRevision = characterFormatRevision
        ' Changing the format.
        run1.CharacterFormat.UnderlineStyle = UnderlineType.Double

        ' 2. Removing the run.
        Dim run2 = paragraph1.Inlines.Cast(Of Run)(1)
        ' Mark run as deleted.
        run2.Revision = New Revision(RevisionType.Delete) With {.Author = "GemBox"}

        ' 3. Inserting a run.
        Dim run3 = New Run(document, "Run3")
        ' Mark run as inserted.
        run3.Revision = New Revision(RevisionType.Insert) With {.Author = "GemBox"}
        paragraph1.Inlines.Add(run3)

        ' 4. Joining paragraphs.
        Dim paragraph2 = section.Blocks.Cast(Of Paragraph)(1)
        ' Marking paragraph as deleted doesn't remove the paragraph content, it joins it with the following paragraph.
        paragraph2.Revision = New Revision(RevisionType.Delete) With {.Author = "GemBox"}
        Dim table = section.Blocks.Cast(Of Table)(3)

        ' 5. Removing a table row.
        Dim row2 = table.Rows(1)
        ' Mark row as deleted.
        row2.Revision = New Revision(RevisionType.Delete) With {.Author = "GemBox"}

        ' 6. Adding a new table row.
        Dim newRow = New TableRow(document, New TableCell(document, New Paragraph(document, "new row")), New TableCell(document))
        ' Mark row as inserted.
        newRow.Revision = New Revision(RevisionType.Insert) With {.Author = "GemBox"}
        table.Rows.Add(newRow)

        document.Save("Added Revisions.docx")
    End Sub

End Module