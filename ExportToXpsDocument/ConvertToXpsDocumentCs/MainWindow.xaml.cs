using System.Windows;
using System.Windows.Controls;
using System.Windows.Xps.Packaging;
using GemBox.Document;

namespace ConvertToXpsDocumentCs
{
    public partial class MainWindow : Window
    {
        XpsDocument xpsDocument;

        public MainWindow()
        {
            InitializeComponent();

            this.SetDocumentViewer(this.DocumentViewer);
        }

        private void SetDocumentViewer(DocumentViewer documentViewer)
        {
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");

            DocumentModel document = new DocumentModel();

            var section = new Section(document);
            document.Sections.Add(section);

            var paragraph = new Paragraph(document);
            section.Blocks.Add(paragraph);

            var run = new Run(document, "Hello World!");
            paragraph.Inlines.Add(run);

            // XpsDocument needs to stay referenced so that DocumentViewer can access additional required resources.
            // Otherwise, GC will collect/dispose XpsDocument and DocumentViewer will not work.
            this.xpsDocument = document.ConvertToXpsDocument(SaveOptions.XpsDefault);      
      
            documentViewer.Document = this.xpsDocument.GetFixedDocumentSequence();
        }
    }
}
