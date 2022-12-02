Imports System.Globalization
Imports GemBox.Document
Imports GemBox.Document.Hyphenation

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Example1()
        Example2()
        Example3()

    End Sub

    Sub Example1()
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
        document.Save("Output1.pdf")
    End Sub

    Sub Example2()
        ' Load Word document from file's path.
        Dim document = DocumentModel.Load("AutoHyphenation.docx")

        ' Enable auto hyphenation.
        document.HyphenationOptions.AutoHyphenation = True

        ' Load hyphenation dictionary from file's path.
        Dim hyphenationDictionary = TexHyphenationDictionary.Load("hyph-en-gb.tex")

        ' Set loaded hyphenation dictionary for specified language.
        DocumentModel.HyphenationDictionaries(New CultureInfo("en-GB")) = hyphenationDictionary

        ' Save output file.
        document.Save("Output2.pdf")
    End Sub

    Sub Example3()
        ' Load Word document from file's path.
        Dim document = DocumentModel.Load("AutoHyphenationML.docx")

        ' Enable auto hyphenation.
        document.HyphenationOptions.AutoHyphenation = True

        ' Load hyphenation dictionaries on demand and set them for specified languages.
        AddHandler DocumentModel.HyphenationDictionaries.HyphenationDictionaryLoading,
            Sub(sender, e)
                Select Case e.CultureInfo.Name

                    Case "en-GB"
                        e.HyphenationDictionary = TexHyphenationDictionary.Load("hyph-en-gb.tex")
                        Exit Select

                    Case "de-DE"
                        e.HyphenationDictionary = TexHyphenationDictionary.Load("hyph-de-1901.tex")
                        Exit Select

                    Case "es-ES"
                        e.HyphenationDictionary = TexHyphenationDictionary.Load("hyph-es.tex")
                        Exit Select
                End Select
            End Sub

        ' Save output file.
        document.Save("Output3.pdf")
    End Sub

End Module