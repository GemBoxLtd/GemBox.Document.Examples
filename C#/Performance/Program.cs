using System;
using System.Collections.Generic;
using System.IO;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Engines;
using BenchmarkDotNet.Jobs;
using BenchmarkDotNet.Running;
using GemBox.Document;

[SimpleJob(RuntimeMoniker.Net48)]
[SimpleJob(RuntimeMoniker.NetCoreApp31)]
public class Program
{
    private DocumentModel document;
    private readonly Consumer consumer = new Consumer();

    public static void Main()
    {
        BenchmarkRunner.Run<Program>();
    }

    [GlobalSetup]
    public void SetLicense()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        // If using Free version and example exceeds its limitations, use Trial or Time Limited version:
        // https://www.gemboxsoftware.com/document/examples/free-trial-professional-modes/1301

        this.document = DocumentModel.Load("RandomSections.docx");
    }

    [Benchmark]
    public DocumentModel Reading()
    {
        return DocumentModel.Load("RandomSections.docx");
    }

    [Benchmark]
    public void Writing()
    {
        using (var stream = new MemoryStream())
            this.document.Save(stream, new DocxSaveOptions());
    }

    [Benchmark]
    public void Iterating()
    {
        this.LoopThroughAllElements().Consume(this.consumer);
    }

    public IEnumerable<Element> LoopThroughAllElements()
    {
        return this.document.GetChildElements(true);
    }
}