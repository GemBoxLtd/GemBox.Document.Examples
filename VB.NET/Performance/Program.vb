Imports System
Imports System.Diagnostics
Imports System.Linq
Imports GemBox.Document

Module Program

    Sub Main()

        ' GemBox.Document has 4 working modes, each mode has the same performance and set of features. 
        ' Read more on: https://www.gemboxsoftware.com/document/help/html/Evaluation_and_Licensing.htm

        ' Set license key to use GemBox.Document in a Free mode.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        ' Continue to use the component in a Trial mode when free limit is reached.
        AddHandler ComponentInfo.FreeLimitReached, Sub(sender, e) e.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial

        Dim watch = Stopwatch.StartNew()
        Dim document = DocumentModel.Load("lorem-ipsum-100-pages.docx", LoadOptions.DocxDefault)
        Console.WriteLine($"Load file [sec]: {watch.Elapsed.TotalSeconds}")

        watch.Restart()
        Dim numberOfParagraphs As Integer = document.GetChildElements(True, ElementType.Paragraph).Count()
        Console.WriteLine($"Iterate through {numberOfParagraphs} paragraphs [sec]: {watch.Elapsed.TotalSeconds}")

        watch.Restart()
        document.Save("output.docx")
        Console.WriteLine($"Save file [sec]: {watch.Elapsed.TotalSeconds}")

    End Sub
End Module