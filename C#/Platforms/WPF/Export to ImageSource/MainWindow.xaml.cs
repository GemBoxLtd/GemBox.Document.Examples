using GemBox.Document;
using System.Windows;

namespace ExportToImageSource
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            this.InitializeComponent();

            // If using the Professional version, put your serial key below.
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");

            this.SetImageControl("Reading.docx", 0);
        }

        private void SetImageControl(string path, int pageIndex)
        {
            var document = DocumentModel.Load(path);

            var imageOptions = new ImageSaveOptions();
            imageOptions.PageNumber = pageIndex;

            var imageSource = document.ConvertToImageSource(imageOptions);
            this.ImageControl.Source = imageSource;
        }
    }
}