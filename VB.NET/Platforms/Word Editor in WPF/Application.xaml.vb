Imports System.Windows.Threading

Public Class Application

    Private Sub OnUnhandledException(sender As Object, e As DispatcherUnhandledExceptionEventArgs)

        MessageBox.Show(Me.MainWindow, e.Exception.ToString(), e.Exception.GetType().ToString(), MessageBoxButton.OK, MessageBoxImage.Error)
        e.Handled = True

    End Sub

End Class
