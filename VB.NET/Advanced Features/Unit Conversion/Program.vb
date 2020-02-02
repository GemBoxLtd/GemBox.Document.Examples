Imports System
Imports GemBox.Document

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document = DocumentModel.Load("Reading.docx")
        Dim pageSetup = document.Sections(0).PageSetup

        Dim widthInPoints As Double = pageSetup.PageWidth
        Dim heightInPoints As Double = pageSetup.PageHeight

        Console.WriteLine("Document's page size in different units:")

        For Each unit As LengthUnit In [Enum].GetValues(GetType(LengthUnit))

            Dim convertedWidth As Double = LengthUnitConverter.Convert(widthInPoints, LengthUnit.Point, unit)
            Dim convertedHeight As Double = LengthUnitConverter.Convert(heightInPoints, LengthUnit.Point, unit)
            Console.WriteLine($"{convertedWidth} x {convertedHeight} {unit.ToString().ToLowerInvariant()}")

        Next

    End Sub
End Module