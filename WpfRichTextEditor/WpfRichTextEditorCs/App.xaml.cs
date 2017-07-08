using System.Windows;
using System.Windows.Threading;

namespace WpfRichTextEditor
{
    public partial class App : Application
    {
        private void OnUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
        {
            MessageBox.Show(this.MainWindow, e.Exception.ToString(), e.Exception.GetType().ToString(), MessageBoxButton.OK, MessageBoxImage.Error);
            e.Handled = true;
        }
    }
}
