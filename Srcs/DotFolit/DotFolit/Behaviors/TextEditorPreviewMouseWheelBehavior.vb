Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Interactivity
Imports ICSharpCode.AvalonEdit

Public Class TextEditorPreviewMouseWheelBehavior
    Inherits Behavior(Of TextEditor)

    Protected Overrides Sub OnAttached()

        MyBase.OnAttached()
        AddHandler Me.AssociatedObject.PreviewMouseWheel, AddressOf TextEditor_PreviewMouseWheel

    End Sub

    Protected Overrides Sub OnDetaching()

        MyBase.OnDetaching()
        RemoveHandler Me.AssociatedObject.PreviewMouseWheel, AddressOf TextEditor_PreviewMouseWheel

    End Sub

    Private Sub TextEditor_PreviewMouseWheel(sender As Object, e As MouseWheelEventArgs)

        ' [WPF]コントロールキーあるいはシフトキーが押されているかどうかを取得する
        ' https://gist.github.com/pinzolo/3080481

        Dim isDownLeftControlKey = ((Keyboard.GetKeyStates(Key.LeftCtrl) And KeyStates.Down) = KeyStates.Down)
        Dim isDownRightControlKey = ((Keyboard.GetKeyStates(Key.RightCtrl) And KeyStates.Down) = KeyStates.Down)
        Dim isDownControlKey = isDownLeftControlKey OrElse isDownRightControlKey

        If isDownControlKey Then

            If 0 < e.Delta Then
                Me.AssociatedObject.FontSize *= 1.1
            Else
                Me.AssociatedObject.FontSize /= 1.1
            End If

        End If

    End Sub


End Class
