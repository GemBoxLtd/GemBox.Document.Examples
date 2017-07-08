using System.Windows;
using System.Windows.Controls;
using System.Windows.Xps.Packaging;
using GemBox.Document;
using Microsoft.Win32;

namespace PrintCs
{
    public partial class MainWindow : Window
    {
        private DocumentModel document;

        public MainWindow()
        {
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");

            InitializeComponent();

            this.EnableControls();
        }

        private void LoadFileBtn_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Filter = "DOCX files (*.docx, *.docm, *.dotm, *.dotx)|*.docx;*.docm;*.dotm;*.dotx|DOC files (*.doc, *.dot)|*.doc;*.dot|HTML files (*.html, *.htm)|*.html;*.htm|RTF files (*.rtf)|*.rtf|TXT files (*.txt)|*.txt";

            if (fileDialog.ShowDialog() == true)
            {
                this.document = DocumentModel.Load(fileDialog.FileName);

                this.ShowPrintPreview();
                this.EnableControls();
            }
        }

        private void SimplePrint_Click(object sender, RoutedEventArgs e)
        {
            // Print to default printer using default options
            this.document.Print();
        }

        private void AdvancedPrint_Click(object sender, RoutedEventArgs e)
        {
            // We can use PrintDialog for defining print options
            PrintDialog printDialog = new PrintDialog();
            printDialog.UserPageRangeEnabled = true;

            if (printDialog.ShowDialog() == true)
            {
                PrintOptions printOptions = new PrintOptions(printDialog.PrintTicket.GetXmlStream());

                printOptions.FromPage = printDialog.PageRange.PageFrom - 1;
                printOptions.ToPage = printDialog.PageRange.PageTo == 0 ? int.MaxValue : printDialog.PageRange.PageTo - 1;

                this.document.Print(printDialog.PrintQueue.FullName, printOptions);
            }
        }

        // We can use DocumentViewer for print preview (but we don't need).
        private void ShowPrintPreview()
        {
            // XpsDocument needs to stay referenced so that DocumentViewer can access additional required resources.
            // Otherwise, GC will collect/dispose XpsDocument and DocumentViewer will not work.
            XpsDocument xpsDocument = document.ConvertToXpsDocument(SaveOptions.XpsDefault);
            this.DocViewer.Tag = xpsDocument;

            this.DocViewer.Document = xpsDocument.GetFixedDocumentSequence();
        }  

        private void EnableControls()
        {
            bool isEnabled = this.document != null;

            this.DocViewer.IsEnabled = isEnabled;
            this.SimplePrintFileBtn.IsEnabled = isEnabled;
            this.AdvancedPrintFileBtn.IsEnabled = isEnabled;
        }
    }
}
