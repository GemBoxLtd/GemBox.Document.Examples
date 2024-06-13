Imports GemBox.Document
Imports GemBox.Document.Hyphenation
Imports System.Globalization

Module Program

    Sub Main()
        Example1()
        Example2()
        Example3()
    End Sub

    Sub Example1()
        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        ' Load Word document from file's path.
        Dim document = DocumentModel.Load("AutoHyphenation.docx")

        ' Get Word document's hyphenation options.
        Dim hyphenationOptions = document.HyphenationOptions

        ' Enable auto hyphenation, set the maximum number of consecutive lines,
        ' and disable hyphenation of words in all capital letters.
        hyphenationOptions.AutoHyphenation = True
        hyphenationOptions.ConsecutiveHyphensLimit = 3
        hyphenationOptions.HyphenateCaps = False

        ' Save output file.
        document.Save("Output.docx")
    End Sub

    Sub Example2()
        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        ' Load Word document from file's path.
        Dim document = DocumentModel.Load("AutoHyphenation.docx")

        ' Enable auto hyphenation.
        document.HyphenationOptions.AutoHyphenation = True

        ' Load hyphenation dictionary from file's path.
        Dim hyphenationDictionary = TexHyphenationDictionary.Load("HyphDictEnGb.tex")

        ' Set loaded hyphenation dictionary for specified language.
        DocumentModel.HyphenationDictionaries(New CultureInfo("en-GB")) = hyphenationDictionary

        ' Save output file.
        document.Save("OutputCustomHyphenation.pdf")
    End Sub

    Sub Example3()
        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        ' Load Word document from file's path.
        Dim document = DocumentModel.Load("AutoHyphenationML.docx")

        ' Enable auto hyphenation.
        document.HyphenationOptions.AutoHyphenation = True

        ' Load hyphenation dictionaries on demand and set them for specified languages.
        AddHandler DocumentModel.HyphenationDictionaries.HyphenationDictionaryLoading,
            Sub(sender, e)
                Select Case e.CultureInfo.Name

                    Case "en-GB"
                        e.HyphenationDictionary = TexHyphenationDictionary.Load("HyphDictEnGb.tex")
                        Exit Select

                    Case "de-DE"
                        e.HyphenationDictionary = TexHyphenationDictionary.Load("HyphDictDe.tex")
                        Exit Select

                    Case "es-ES"
                        e.HyphenationDictionary = TexHyphenationDictionary.Load("HyphDictEs.tex")
                        Exit Select
                End Select
            End Sub

        ' Save output file.
        document.Save("OutputMultiLanguage.pdf")
    End Sub

End Module
