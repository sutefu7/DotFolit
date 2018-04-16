Imports Livet

Class Application

    ' Startup、Exit、DispatcherUnhandledException などのアプリケーション レベルのイベントは、
    ' このファイルで処理できます。

    Private Sub Application_Startup(sender As Object, e As StartupEventArgs)

        DispatcherHelper.UIDispatcher = Dispatcher

    End Sub


End Class
