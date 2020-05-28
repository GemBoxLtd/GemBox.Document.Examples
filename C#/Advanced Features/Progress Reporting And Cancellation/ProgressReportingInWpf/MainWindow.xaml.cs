using System.Windows;
using System.Threading;
using System.Threading.Tasks;
using GemBox.Document;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        // If using Professional version, put your serial key below
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");
        // Use Trial Mode
        ComponentInfo.FreeLimitReached += (eventSender, args) => args.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial;
        InitializeComponent();
    }

    private async void loadButton_Click(object sender, RoutedEventArgs e)
    {
        // Capture the current context on the UI thread
        var context = SynchronizationContext.Current;

        // Create load options
        var loadOptions = new DocxLoadOptions();
        loadOptions.ProgressChanged += (eventSender, args) =>
        {
            var percentage = args.ProgressPercentage;
            // Invoke on the UI thread
            context.Post(progressPercentage =>
            {
                // Update UI
                this.progressBar.Value = (int)progressPercentage;
                this.percentageLabel.Content = progressPercentage.ToString() + "%";
            }, percentage);
        };

        this.percentageLabel.Content = "0%";
        // Use tasks to run the load operation in a new thread
        var file = await Task.Run(() => DocumentModel.Load("LargeDocument.docx", loadOptions));
    }
}