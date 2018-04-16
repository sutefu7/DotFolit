Public Class SolutionExplorerView

    ' ノードを左クリックで選択した後に右クリックしないで、直接右クリック→コンテキストメニューをしようとした場合、
    ' ノードが未選択状態で、コンテキストメニューが表示されてしまう。
    ' コンテキストメニュー自体は、正しく対象ノードを判定できているが、見た目的には、ノード未選択状態でのコンテキストメニュー表示は気に入らない（自然ではない）
    ' 
    ' 対策は、以下のサイトの通りとした
    '
    ' TreeViewで右クリックしたTreeViewItemを選択状態にしたい
    ' https://social.msdn.microsoft.com/Forums/netframework/ja-JP/82c27575-f26c-4ba4-a3d6-f066e21dc91d/treeviewtreeviewitem?forum=wpfja
    ' 
    ' 以下のサイトを参考にして解決したとのこと
    ' how to make mouse right button(click) select treeviewitem?
    ' https://social.msdn.microsoft.com/Forums/vstudio/en-US/b980bac3-3fa6-4a84-b572-e53ce28c64f3/how-to-make-mouse-right-buttonclick-select-treeviewitem?forum=wpf
    '

    Private Sub TreeViewItem_MouseRightButtonDown(sender As Object, e As MouseButtonEventArgs)

        Dim item = TryCast(sender, TreeViewItem)
        If item Is Nothing Then
            Return
        End If

        item.IsSelected = True
        e.Handled = True

    End Sub

End Class
