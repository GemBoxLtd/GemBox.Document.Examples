Imports GemBox.Document
Imports GemBox.Document.Vba

Module Program

    Sub Main()

        Example1()
        Example2()

    End Sub

    Sub Example1()
        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document As New DocumentModel()

        Dim section As New Section(document)
        document.Sections.Add(section)

        ' Create the module.
        Dim vbaModule As VbaModule = document.VbaProject.Modules.Add("SampleModule", VbaModuleType.Document)
        vbaModule.Code =
"Sub WriteHello()
    Selection.TypeText Text:=""Hello World!""
End Sub"

        ' Save the document as macro-enabled Word file.
        document.Save("AddVbaModule.docm")
    End Sub

    Sub Example2()
        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim document = DocumentModel.Load("SampleVba.docm")

        ' Get the module.
        Dim vbaModule As VbaModule = document.VbaProject.Modules("ThisDocument")
        ' Update text for the popup message.
        vbaModule.Code = vbaModule.Code.Replace("Hello world!", "Hello from GemBox.Document!")

        document.Save("UpdateVbaModule.docm")
    End Sub
End Module
