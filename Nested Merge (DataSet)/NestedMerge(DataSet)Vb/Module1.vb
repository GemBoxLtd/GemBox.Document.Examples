Imports System
Imports System.Data
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

        ' Create data source
        Dim ds As New DataSet()

        ' Project details
        Dim projects As New DataTable("Projects")
        projects.Columns.Add("ProjectId", GetType(Integer))
        projects.Columns.Add("ContactName", GetType(String))
        projects.Columns.Add("ProjectName", GetType(String))
        ds.Tables.Add(projects)

        ' Item details
        Dim items As New DataTable("Items")
        items.Columns.Add("ProjectId", GetType(Integer))
        items.Columns.Add("Date", GetType(DateTime))
        items.Columns.Add("Hours", GetType(Double))
        items.Columns.Add("Price", GetType(Double))
        items.Columns.Add("Total", GetType(Double))
        ds.Tables.Add(items)

        ' Add parent-child relation 
        ds.Relations.Add("Items", projects.Columns("ProjectId"), items.Columns("ProjectId"))

        ' Fill DataSource.
        For i As Integer = 0 To numberOfProjects - 1
            Dim projectId As Integer = i + 1
            Dim contactName As String = "John Doe"
            Dim projectName As String = "Project " & projectId
            projects.Rows.Add(projectId, contactName, projectName)

            Dim startDate As DateTime = DateTime.Now.AddDays(-itemsPerProject)
            For j As Integer = 0 To itemsPerProject - 1
                ' We worked between 6 and 8 hours per day.
                Dim workHours As Integer = ((i + 1) * j) Mod 3 + 6
                Dim totalPrice As Double = workHours * 35
                items.Rows.Add(projectId, startDate.AddDays(j), workHours, 35, totalPrice)
            Next
        Next

        AddHandler document.MailMerge.FieldMerging,
            Sub(sender, e)
                If e.IsValueFound Then
                    ' Define custom formatting.
                    Select Case e.FieldName
                        Case "Date"
                            DirectCast(e.Inline, Run).Text = DirectCast(e.Value, DateTime).ToString("dddd, MMMM d, yyyy")
                            Exit Select
                        Case "Price", "Total"
                            DirectCast(e.Inline, Run).Text = CDbl(e.Value).ToString("0.00")
                            Exit Select
                    End Select
                ElseIf e.RangeName = "Projects" AndAlso e.FieldName = "TotalPrice" Then
                    ' For each project calculate Total.
                    e.Inline = New Run(e.Document, ds.Tables(e.RangeName).Rows(e.RecordNumber - 1).GetChildRows("Items").Sum(Function(item) CDbl(item("Total"))).ToString("0.00"))
                    e.Cancel = False
                End If

            End Sub

        ' Execute nested mail merge.
        document.MailMerge.Execute(ds, Nothing)

        document.Save("Nested Merge (DataSet).docx")

    End Sub

End Module