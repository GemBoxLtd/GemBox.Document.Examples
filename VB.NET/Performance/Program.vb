Imports BenchmarkDotNet.Attributes
Imports BenchmarkDotNet.Engines
Imports BenchmarkDotNet.Jobs
Imports BenchmarkDotNet.Running
Imports GemBox.Document
Imports System.Collections.Generic
Imports System.IO

<SimpleJob(RuntimeMoniker.Net80)>
<SimpleJob(RuntimeMoniker.Net48)>
Public Class Program

    Private document As DocumentModel
    Private ReadOnly consumer As Consumer = New Consumer()

    Public Shared Sub Main()
        BenchmarkRunner.Run(Of Program)()
    End Sub

    <GlobalSetup>
    Public Sub SetLicense()
        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        ' If using Free version and example exceeds its limitations, use Trial or Time Limited version:
        ' https://www.gemboxsoftware.com/document/examples/free-trial-professional/1301

        Me.document = DocumentModel.Load("RandomSections.docx")
    End Sub

    <Benchmark>
    Public Function Reading() As DocumentModel
        Return DocumentModel.Load("RandomSections.docx")
    End Function

    <Benchmark>
    Public Sub Writing()
        Using stream = New MemoryStream()
            Me.document.Save(stream, New DocxSaveOptions())
        End Using
    End Sub

    <Benchmark>
    Public Sub Iterating()
        Me.LoopThroughAllElements().Consume(Me.consumer)
    End Sub

    Public Function LoopThroughAllElements() As IEnumerable(Of Element)
        Return Me.document.GetChildElements(True)
    End Function

End Class
