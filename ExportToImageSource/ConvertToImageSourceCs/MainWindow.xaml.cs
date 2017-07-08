using System.Windows;
using System.Windows.Controls;
using GemBox.Document;

namespace ConvertToImageSourceCs
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            SetImageSource(this.ImageControl);
        }

        private static void SetImageSource(Image image)
        {
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");

            DocumentModel document = new DocumentModel();

            var section = new Section(document);
            document.Sections.Add(section);

            var paragraph = new Paragraph(document);
            section.Blocks.Add(paragraph);

            var run = new Run(document, "Hello World!");
            paragraph.Inlines.Add(run);

            image.Source = document.ConvertToImageSource(SaveOptions.ImageDefault);
        }
    }
}
