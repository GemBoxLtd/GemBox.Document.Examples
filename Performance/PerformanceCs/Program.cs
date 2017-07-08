using System;
using System.Diagnostics;
using GemBox.Document;

namespace PerformanceCs
{
    class Program
    {
        static void Main(string[] args)
        {
            // If using Professional version, put your serial key below.
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");

            // If sample exceeds Free version limitations then continue as trial version: 
            // https://www.gemboxsoftware.com/Document/help/html/Evaluation_and_Licensing.htm
            ComponentInfo.FreeLimitReached += (sender, e) => e.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial;

            Console.WriteLine("Performance sample:");
            Console.WriteLine();

            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();

            DocumentModel document = DocumentModel.Load("Template.docx", LoadOptions.DocxDefault);

            Console.WriteLine("Load file (seconds): " + stopwatch.Elapsed.TotalSeconds);

            stopwatch.Reset();
            stopwatch.Start();

            int numberOfParagraphs = 0;
            foreach (var item in document.GetChildElements(true, ElementType.Paragraph))
                ++numberOfParagraphs;

            Console.WriteLine("Iterate through " + numberOfParagraphs + " paragraphs (seconds): " + stopwatch.Elapsed.TotalSeconds);

            stopwatch.Reset();
            stopwatch.Start();

            document.Save("Report.docx");

            Console.WriteLine("Save file (seconds): " + stopwatch.Elapsed.TotalSeconds);
        }
    }
}
