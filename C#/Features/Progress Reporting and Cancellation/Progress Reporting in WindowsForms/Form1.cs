using GemBox.Document;
using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

public partial class MainForm : Form
{
    public MainForm()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");
        InitializeComponent();
    }

    private async void loadButton_Click(object sender, EventArgs e)
    {
        // Capture the current context on the UI thread.
        var context = SynchronizationContext.Current;

        // Create load options.
        var loadOptions = new DocxLoadOptions();
        loadOptions.ProgressChanged += (eventSender, args) =>
        {
            var percentage = args.ProgressPercentage;
            // Invoke on the UI thread.
            context.Post(progressPercentage =>
            {
                // Update UI.
                this.progressBar.Value = (int)progressPercentage;
                this.percentageLabel.Text = progressPercentage.ToString() + "%";
            }, percentage);
        };

        this.percentageLabel.Text = "0%";
        // Use tasks to run the load operation in a new thread.
        var file = await Task.Run(() => DocumentModel.Load("LargeDocument.docx", loadOptions));
    }
}
