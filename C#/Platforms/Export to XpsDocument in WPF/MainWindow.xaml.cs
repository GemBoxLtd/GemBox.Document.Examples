using GemBox.Document;
using System.Windows;
using System.Windows.Xps.Packaging; // Add reference for 'ReachFramework'.

namespace ExportToXpsDocument
{
    public partial class MainWindow : Window
    {
        private XpsDocument xpsDocument;

        public MainWindow()
        {
            this.InitializeComponent();

            // If using the Professional version, put your serial key below.
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");

            this.SetDocumentViewer("Reading.docx");
        }

        private void SetDocumentViewer(string path)
        {
            var document = DocumentModel.Load(path);

            // XpsDocument object must stay referenced so that DocumentViewer can access its resources.
            // Otherwise, GC will collect/dispose XpsDocument and DocumentViewer will no longer work.
            this.xpsDocument = document.ConvertToXpsDocument(SaveOptions.XpsDefault);

            this.DocumentViewer.Document = this.xpsDocument.GetFixedDocumentSequence();
        }
    }
}