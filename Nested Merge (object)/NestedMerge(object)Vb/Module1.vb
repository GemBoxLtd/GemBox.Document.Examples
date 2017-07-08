Imports System
Imports System.Linq
Imports GemBox.Document

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As DocumentModel = DocumentModel.Load("InvoiceForNestedMailMerge.docx")

        Dim numberOfProjects As Integer = 4
        Dim itemsPerProject As Integer = 7

        ' Fill document header.
        document.MailMerge.Execute(New With {
            .CompanyName = "ACME Corporation",
            .Address = "240 Old Country Road, Springfield, IL",
            .PrintDate = DateTime.Now.ToLongDateString()
        })

        ' Create data source using LINQ and anonymous types.
        Dim projects = Enumerable.Range(0, numberOfProjects).Select(Function(projectIndex) New With {
          .ProjectId = projectIndex + 1,
          .ContactName = "John Doe",
          .ProjectName = "Project " & (projectIndex + 1),
          .Items = Enumerable.Range(0, itemsPerProject).Select(Function(itemIndex) New With {
            .Date = DateTime.Now.AddDays(-itemsPerProject + itemIndex),
            .Hours = ((projectIndex + 1) * itemIndex) Mod 3 + 6,
            .Price = 35,
            .Total = (((projectIndex + 1) * itemIndex) Mod 3 + 6) * 35
            }).ToArray()
        }).ToArray()

        AddHandler document.MailMerge.FieldMerging,
            Sub(sender, e)
                If e.IsValueFound Then
                    ' Define custom formatting.
                    Select Case e.FieldName
                        Case "Date"
                            DirectCast(e.Inline, Run).Text = DirectCast(e.Value, DateTime).ToString("dddd, MMMM d, yyyy")
                            Exit Select
                        Case "Price", "Total"
                            DirectCast(e.Inline, Run).Text = CInt(e.Value).ToString("0.00")
                            Exit Select
                    End Select
                ElseIf e.RangeName = "Projects" AndAlso e.FieldName = "TotalPrice" Then
                    ' For each project calculate Total.
                    e.Inline = New Run(e.Document, projects(e.RecordNumber - 1).Items.Sum(Function(item) item.Total).ToString("0.00"))
                    e.Cancel = False
                End If
            End Sub

        ' Execute nested mail merge.
        document.MailMerge.Execute(projects, "Projects")

        document.Save("Nested Merge (object).docx")

    End Sub

End Module