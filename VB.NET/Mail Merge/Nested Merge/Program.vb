Imports GemBox.Document
Imports System
Imports System.Data
Imports System.Linq

Module Program

    Sub Main()
        Example1()
        Example2()
    End Sub

    Sub Example1()
        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim numberOfProjects As Integer = 3
        Dim itemsPerProject As Integer = 8

        Dim projectsRangeName As String = "Projects"
        Dim itemsRangeName As String = "Items"

        ' Create relational data.
        Dim projects As New DataTable(projectsRangeName)
        projects.Columns.Add("Id", GetType(Integer))
        projects.Columns.Add("Name", GetType(String))

        Dim items As New DataTable(itemsRangeName)
        items.Columns.Add("ProjectId", GetType(Integer))
        items.Columns.Add("Date", GetType(DateTime))
        items.Columns.Add("Hours", GetType(Integer))
        items.Columns.Add("Unit", GetType(Double))
        items.Columns.Add("Price", GetType(Double))

        ' Create DataSet with parent-child relation.
        Dim data As New DataSet()
        data.Tables.Add(projects)
        data.Tables.Add(items)
        data.Relations.Add(itemsRangeName, projects.Columns("Id"), items.Columns("ProjectId"))

        For projectIndex As Integer = 1 To numberOfProjects

            Dim id As Integer = projectIndex
            Dim name As String = $"Project {projectIndex}"

            projects.Rows.Add(id, name)

            For itemIndex As Integer = 1 To itemsPerProject

                Dim [date] As DateTime = DateTime.Today _
                    .AddMonths(projectIndex - numberOfProjects) _
                    .AddDays(itemIndex - itemsPerProject)
                Dim hours As Integer = itemIndex Mod 3 + 6
                Dim unit As Double = projectIndex * 35.0
                Dim price As Double = hours * unit

                items.Rows.Add(id, [date], hours, unit, price)

            Next
        Next

        Dim document = DocumentModel.Load("MergeNestedRanges.docx")

        ' Customize mail merging to achieve calculation of "TotalPrice" for each project.
        AddHandler document.MailMerge.FieldMerging,
            Sub(sender, e)
                If e.MergeContext.RangeName = "Projects" And e.FieldName = "TotalPrice" Then

                    Dim total = data.Tables(e.MergeContext.RangeName).Rows(e.MergeContext.RecordIndex) _
                        .GetChildRows(itemsRangeName).Sum(Function(item) CDbl(item("Price")))

                    Dim totalRun As New Run(e.Document, total.ToString("0.00"))
                    totalRun.CharacterFormat = e.Field.CharacterFormat.Clone()

                    e.Inline = totalRun
                    e.Cancel = False

                End If
            End Sub

        ' Execute nested mail merge.
        document.MailMerge.Execute(data, Nothing)

        document.Save("MergedNestedRangesOutput.docx")
    End Sub

    Sub Example2()
        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim numberOfProjects As Integer = 3
        Dim itemsPerProject As Integer = 8

        ' Create hierarchical data source using LINQ and anonymous types.
        Dim projects = Enumerable.Range(1, numberOfProjects).Select(
            Function(projectIndex)
                Return New With
                {
                    .Name = $"Project {projectIndex}",
                    .Items = Enumerable.Range(1, itemsPerProject).Select(
                        Function(itemIndex)

                            Dim [date] As DateTime = DateTime.Today _
                                .AddMonths(projectIndex - numberOfProjects) _
                                .AddDays(itemIndex - itemsPerProject)
                            Dim hours As Integer = itemIndex Mod 3 + 6
                            Dim unit As Double = projectIndex * 35.0
                            Dim price As Double = hours * unit

                            Return New With {.Date = [date], .Hours = hours, .Unit = unit, .Price = price}
                        End Function).ToArray()
                }
            End Function).ToArray()

        Dim document = DocumentModel.Load("MergeNestedRanges.docx")

        ' Customize mail merging to achieve calculation of "TotalPrice" for each project.
        AddHandler document.MailMerge.FieldMerging,
            Sub(sender, e)
                If e.MergeContext.RangeName = "Projects" And e.FieldName = "TotalPrice" Then

                    Dim total = projects(e.MergeContext.RecordIndex).Items.Sum(Function(item) item.Price)

                    Dim totalRun As New Run(e.Document, total.ToString("0.00"))
                    totalRun.CharacterFormat = e.Field.CharacterFormat.Clone()

                    e.Inline = totalRun
                    e.Cancel = False

                End If
            End Sub

        ' Execute nested mail merge.
        document.MailMerge.Execute(projects, "Projects")

        document.Save("MergedNestedRangesOutputWithObject.docx")
    End Sub

End Module
