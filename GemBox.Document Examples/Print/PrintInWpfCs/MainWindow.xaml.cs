using System.Windows;
using System.Windows.Controls;
using System.Windows.Xps.Packaging;
using Microsoft.Win32;
using GemBox.Document;

public partial class MainWindow : Window
{
    private DocumentModel document;

    public MainWindow()
    {
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");
        InitializeComponent();
    }

    private void LoadFileBtn_Click(object sender, RoutedEventArgs e)
    {
        OpenFileDialog openFileDialog = new OpenFileDialog();
        openFileDialog.Filter =
            "DOCX files (*.docx, *.dotx, *.docm, *.dotm)|*.docx;*.dotx;*.docm;*.dotm" +
            "|DOC files (*.doc, *.dot)|*.doc;*.dot" +
            "|RTF files (*.rtf)|*.rtf" +
            "|HTML files (*.html, *.htm)|*.html;*.htm" +
            "|PDF files (*.pdf)|*.pdf" +
            "|TXT files (*.txt)|*.txt";

        if (openFileDialog.ShowDialog() == true)
        {
            this.document = DocumentModel.Load(openFileDialog.FileName);
            this.ShowPrintPreview();
        }
    }

    private void PrintFileBtn_Click(object sender, RoutedEventArgs e)
    {
        if (this.document == null)
            return;

        PrintDialog printDialog = new PrintDialog() { UserPageRangeEnabled = true };
        if (printDialog.ShowDialog() == true)
        {
            PrintOptions printOptions = new PrintOptions(printDialog.PrintTicket.GetXmlStream());

            printOptions.FromPage = printDialog.PageRange.PageFrom - 1;
            printOptions.ToPage = printDialog.PageRange.PageTo == 0 ? int.MaxValue : printDialog.PageRange.PageTo - 1;

            this.document.Print(printDialog.PrintQueue.FullName, printOptions);
        }
    }

    private void ShowPrintPreview()
    {
        XpsDocument xpsDocument = this.document.ConvertToXpsDocument(SaveOptions.XpsDefault);

        // Note, XpsDocument must stay referenced so that DocumentViewer can access additional resources from it.
        // Otherwise, GC will collect/dispose XpsDocument and DocumentViewer will no longer work.
        this.DocViewer.Tag = xpsDocument;
        this.DocViewer.Document = xpsDocument.GetFixedDocumentSequence();
    }
}