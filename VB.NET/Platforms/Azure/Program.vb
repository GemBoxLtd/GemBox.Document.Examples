Imports Microsoft.Extensions.Hosting

Public Class Program

    Public Shared Sub Main()
        Dim host As IHost = New HostBuilder() _
            .ConfigureFunctionsWorkerDefaults() _
            .Build()
        host.Run()
    End Sub

End Class
