using System.Windows;
using GemBox.Document;

namespace ExportToImageSource
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            this.InitializeComponent();

            // If using the Professional version, put your serial key below.
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");

            string path = "Reading.docx";
            int pageIndex = 0;

            this.SetImageControl(path, pageIndex);
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