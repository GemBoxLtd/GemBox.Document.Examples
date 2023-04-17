Imports GemBox.Document

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        ' If using the Professional version, remove this FreeLimitReached event handler.
        AddHandler ComponentInfo.FreeLimitReached, Sub(sender, e) e.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial

        Example1()
        Example2()

    End Sub

    Sub Example1()
        ' Word files that will be combined into one file.
        Dim files As String() =
        {
            "MergeFile01.docx",
            "MergeFile02.docx",
            "MergeFile03.docx"
        }

        ' Create destination document.
        Dim destination As New DocumentModel()

        ' Merge multiple source documents by importing their content at the end.
        For Each file In files
            Dim source = DocumentModel.Load(file)
            destination.Content.End.InsertRange(source.Content)
        Next

        ' Save joined documents into one file.
        destination.Save("Merged Files.docx")
    End Sub

    Sub Example2()
        ' Word files that will be combined into one file.
        Dim files As String() =
        {
            "MergeFile01.docx",
            "MergeFile02.docx",
            "MergeFile03.docx"
        }

        Dim destination As New DocumentModel()
        Dim firstSourceDocument = True

        For Each file In files

            Dim source = DocumentModel.Load(file)
            Dim firstSourceSection = True

            ' Reuse the same mapping for importing to improve performance.
            Dim mapping = New ImportMapping(source, destination, False)

            For Each sourceSection In source.Sections

                ' Import section from source document to destination document.
                Dim destinationSection = destination.Import(sourceSection, True, mapping)
                destination.Sections.Add(destinationSection)

                ' Set the first section to start on the same page as the previous section.
                ' In other words, the source content continues to flow with the current destination content.
                If firstSourceSection Then
                    destinationSection.PageSetup.SectionStart = SectionStart.Continuous
                    firstSourceSection = False
                End If
            Next

            ' Set the destination's default formatting to first source's default formatting.
            ' Note, a single document can only have one default formatting.
            If firstSourceDocument Then
                destination.DefaultCharacterFormat = source.DefaultCharacterFormat.Clone()
                destination.DefaultParagraphFormat = source.DefaultParagraphFormat.Clone()
                firstSourceDocument = False
            End If
        Next

        ' Save joined sections into one file.
        destination.Save("Merged Sections.docx")
    End Sub

End Module