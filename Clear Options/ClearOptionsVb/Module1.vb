Imports System
Imports GemBox.Document
Imports GemBox.Document.MailMerging

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As DocumentModel = DocumentModel.Load("MailMergeClearOptions.docx")

        ' Example 1: Data source will remove "First choice" paragraph because there is no value defined for FirstChoice field.
        Dim dataSource1 = New With {
            .Header = "My header",
            .SecondChoice = "I have chosen second choice."
        }

        document.MailMerge.ClearOptions = MailMergeClearOptions.RemoveEmptyParagraphs
        document.MailMerge.Execute(dataSource1, "Example1")

        ' Example 2: Data source will remove table row with label "Address" because value for field Address is Nothing.
        Dim dataSource2 = New With {
            .Name = "John Doe",
            .Email = "john.doe@acme.com",
            .Address = DirectCast(Nothing, String),
            .Country = "USA"
        }

        document.MailMerge.ClearOptions = MailMergeClearOptions.RemoveEmptyTableRows
        document.MailMerge.Execute(dataSource2, "Example2")

        ' Example 3: Data source will remove mail merge range for second item because it has both Header and Content values that are string.Empty or Nothing.
        Dim dataSource3 = New With {
            .Count = 2,
            .HeaderContent = New Object() {New With {
                .Header = "My header 1",
                .Content = "My content 1." & vbLf & "Second line of my content 1."
            }, New With {
                .Header = String.Empty,
                .Content = DirectCast(Nothing, Object)
            }, New With {
                .Header = "My header 3",
                .Content = "My content 3." & vbLf & "Second line of my content 3."
            }}
        }
        document.MailMerge.ClearOptions = MailMergeClearOptions.RemoveEmptyRanges Or MailMergeClearOptions.RemoveEmptyParagraphs
        document.MailMerge.Execute(dataSource3, "Example3")

        document.Save("Clear Options.docx")

    End Sub

End Module