using GemBox.Document;
using GemBox.Document.Hyphenation;
using System.Globalization;

class Program
{
    static void Main()
    {
        Example1();
        Example2();
        Example3();
    }

    static void Example1()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        // Load Word document from file's path.
        var document = DocumentModel.Load("AutoHyphenation.docx");

        // Get Word document's hyphenation options.
        var hyphenationOptions = document.HyphenationOptions;

        // Enable auto hyphenation, set the maximum number of consecutive lines,
        // and disable hyphenation of words in all capital letters.
        hyphenationOptions.AutoHyphenation = true;
        hyphenationOptions.ConsecutiveHyphensLimit = 3;
        hyphenationOptions.HyphenateCaps = false;

        // Save output file.
        document.Save("Output.docx");
    }

    static void Example2()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        // Load Word document from file's path.
        var document = DocumentModel.Load("AutoHyphenation.docx");

        // Enable auto hyphenation.
        document.HyphenationOptions.AutoHyphenation = true;

        // Load hyphenation dictionary from file's path.
        var hyphenationDictionary = TexHyphenationDictionary.Load("HyphDictEnGb.tex");

        // Set loaded hyphenation dictionary for specified language.
        DocumentModel.HyphenationDictionaries[new CultureInfo("en-GB")] = hyphenationDictionary;

        // Save output file.
        document.Save("OutputCustomHyphenation.pdf");
    }

    static void Example3()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        // Load Word document from file's path.
        var document = DocumentModel.Load("AutoHyphenationML.docx");

        // Enable auto hyphenation.
        document.HyphenationOptions.AutoHyphenation = true;

        // Load hyphenation dictionaries on demand and set them for specified languages.
        DocumentModel.HyphenationDictionaries.HyphenationDictionaryLoading +=
            (sender, e) =>
            {
                switch (e.CultureInfo.Name)
                {
                    case "en-GB":
                        e.HyphenationDictionary = TexHyphenationDictionary.Load("HyphDictEnGb.tex");
                        break;
                    case "de-DE":
                        e.HyphenationDictionary = TexHyphenationDictionary.Load("HyphDictDe.tex");
                        break;
                    case "es-ES":
                        e.HyphenationDictionary = TexHyphenationDictionary.Load("HyphDictEs.tex");
                        break;
                }
            };

        // Save output file.
        document.Save("OutputMultiLanguage.pdf");
    }
}
