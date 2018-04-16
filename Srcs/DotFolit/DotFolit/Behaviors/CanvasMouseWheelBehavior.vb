Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Interactivity

Public Class CanvasMouseWheelBehavior
    Inherits Behavior(Of Canvas)

    Protected Overrides Sub OnAttached()

        MyBase.OnAttached()
        AddHandler Me.AssociatedObject.MouseWheel, AddressOf Canvas_MouseWheel

    End Sub

    Protected Overrides Sub OnDetaching()

        MyBase.OnDetaching()
        RemoveHandler Me.AssociatedObject.MouseWheel, AddressOf Canvas_MouseWheel

    End Sub

    Private Sub Canvas_MouseWheel(sender As Object, e As MouseWheelEventArgs)

        ' [WPF]コントロールキーあるいはシフトキーが押されているかどうかを取得する
        ' https://gist.github.com/pinzolo/3080481

        Dim isDownLeftControlKey = ((Keyboard.GetKeyStates(Key.LeftCtrl) And KeyStates.Down) = KeyStates.Down)
        Dim isDownRightControlKey = ((Keyboard.GetKeyStates(Key.RightCtrl) And KeyStates.Down) = KeyStates.Down)
        Dim isDownControlKey = isDownLeftControlKey OrElse isDownRightControlKey

        If isDownControlKey Then

            Dim newScale = TryCast(Me.AssociatedObject.RenderTransform, ScaleTransform)
            If newScale Is Nothing Then
                newScale = New ScaleTransform With {.ScaleX = 1, .ScaleY = 1}
                Me.AssociatedObject.RenderTransform = newScale
            End If

            If 0 < e.Delta Then

                newScale.ScaleX *= 1.1
                newScale.ScaleY *= 1.1

            Else

                newScale.ScaleX /= 1.1
                newScale.ScaleY /= 1.1

            End If

        End If

    End Sub

End Class
