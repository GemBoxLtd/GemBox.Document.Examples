Imports System
Imports System.Text
Imports GemBox.Document

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As DocumentModel = DocumentModel.Load("Reading.docx")

        Dim sb As New StringBuilder()

        sb.AppendLine("Page size (width X height):")

        Dim width As Double = document.Sections(0).PageSetup.PageWidth
        Dim height As Double = document.Sections(0).PageSetup.PageHeight

        For Each unit As LengthUnit In [Enum].GetValues(GetType(LengthUnit))

            sb.AppendFormat("{0} X {1} {2}",
                            LengthUnitConverter.Convert(width, LengthUnit.Point, unit),
                            LengthUnitConverter.Convert(height, LengthUnit.Point, unit),
                            unit.ToString().ToLowerInvariant())

            sb.AppendLine()

        Next

        Console.WriteLine(sb.ToString())

    End Sub

End Module