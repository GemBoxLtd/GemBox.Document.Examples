using GemBox.Document;

namespace DocumentMaui
{
    public partial class MainPage : ContentPage
    {
        static MainPage()
        {
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");
        }
        public MainPage()
        {
            InitializeComponent();
        }
        private async Task<string> CreateDocument()
        {
            var document = new DocumentModel();

            document.Sections.Add(new Section(document,
                new Paragraph(document, text.Text),
                new Paragraph(document, new Picture(document, await FileSystem.OpenAppPackageFileAsync("fragonard_reader.jpg")))));

            var filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Example.pdf");

            document.Save(filePath);

            return filePath;
        }

        private async void Button_Clicked(object sender, EventArgs e)
        {
            button.IsEnabled = false;
            activity.IsRunning = true;

            var filePath = await CreateDocument();
            await Launcher.OpenAsync(new OpenFileRequest(Path.GetFileName(filePath), new ReadOnlyFile(filePath)));

            activity.IsRunning = false;
            button.IsEnabled = true;
        }
    }
}