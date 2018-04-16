Imports Livet
Imports Livet.Messaging.Windows

Public Class ProgressViewModel
    Inherits ViewModel

    Public Property DoWork As Action = Nothing
    Public Property IsDoWorkFinishedAfterWait As Boolean = False

    Public Async Sub Initialize()

        ' 受け取った処理を実行
        Await Task.Run(Sub() DoWork.Invoke())

        ' 処理は終わっているが、なかなかツリー表示されない現象の仮対応。ツリー表示されてから進捗画面を閉じるようにする
        '（デバッグで出力ペインを見ると、Roslyn 関係のアセンブリファイルを読み込んでいる？）
        ' 5 秒待機は妥当ではないかもしれない（プロジェクト数や環境の違いによって）
        If IsDoWorkFinishedAfterWait Then
            Await Task.Delay(5000)
        End If

        ' 処理が終わったら、自動で画面終了
        Await Messenger.RaiseAsync(New WindowActionMessage("WindowCloseAction"))

    End Sub

End Class
