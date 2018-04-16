Imports System.Runtime.CompilerServices
Imports Livet
Imports Livet.Messaging


Module ViewModelExtensions

    ' 情報メッセージ
    <Extension()>
    Public Async Function ShowInformationMessageAsync(self As ViewModel, messageBoxText As String, Optional caption As String = "情報") As Task

        Dim message = New InformationMessage(messageBoxText, caption, MessageBoxImage.Information, "Information")
        Await self.Messenger.RaiseAsync(message)

    End Function

    ' 警告メッセージ
    <Extension()>
    Public Async Function ShowWarningMessageAsync(self As ViewModel, messageBoxText As String, Optional caption As String = "警告") As Task

        Dim message = New InformationMessage(messageBoxText, caption, MessageBoxImage.Warning, "Warning")
        Await self.Messenger.RaiseAsync(message)

    End Function

    ' エラーメッセージ
    <Extension()>
    Public Async Function ShowErrorMessageAsync(self As ViewModel, messageBoxText As String, Optional caption As String = "エラー") As Task

        Dim message = New InformationMessage(messageBoxText, caption, MessageBoxImage.Error, "Error")
        Await self.Messenger.RaiseAsync(message)

    End Function

    ' 確認メッセージ(OKCancel)
    <Extension()>
    Public Async Function ShowConfirmationOKCancelMessageAsync(self As ViewModel, messageBoxText As String, Optional caption As String = "確認") As Task(Of Boolean)

        Dim message = New ConfirmationMessage(messageBoxText, caption, MessageBoxImage.Question, MessageBoxButton.OKCancel, "Confirmation")
        Await self.Messenger.RaiseAsync(message)

        Return message.Response.GetValueOrDefault()

    End Function

    ' 確認メッセージ(YesNo)
    <Extension()>
    Public Async Function ShowConfirmationYesNoMessageAsync(self As ViewModel, messageBoxText As String, Optional caption As String = "確認") As Task(Of Boolean)

        Dim message = New ConfirmationMessage(messageBoxText, caption, MessageBoxImage.Question, MessageBoxButton.YesNo, "Confirmation")
        Await self.Messenger.RaiseAsync(message)

        Return message.Response.GetValueOrDefault()

    End Function

End Module
