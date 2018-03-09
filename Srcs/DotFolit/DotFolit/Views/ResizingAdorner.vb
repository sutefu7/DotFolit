Imports System
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives
Imports System.Windows.Documents
Imports System.Windows.Input
Imports System.Windows.Media


' DENIS VUYKA
' WPF. SIMPLE ADORNER USAGE WITH DRAG AND RESIZE OPERATIONS
' https://denisvuyka.wordpress.com/2007/10/15/wpf-simple-adorner-usage-with-drag-and-resize-operations/

' さらに、以下の変更を加えました。
' ・上下左右への Thumb の追加
' ・各 Thumb の表示位置を、より外側へ変更


Public Class ResizingAdorner
    Inherits Adorner

    ' 上、下、左、右（それぞれ線の中央）
    Private TopThumb As Thumb = Nothing
    Private BottomThumb As Thumb = Nothing
    Private RightThumb As Thumb = Nothing
    Private LeftThumb As Thumb = Nothing

    ' 左上、右上、右下、左下
    Private TopLeftThumb As Thumb = Nothing
    Private TopRightThumb As Thumb = Nothing
    Private BottomLeftThumb As Thumb = Nothing
    Private BottomRightThumb As Thumb = Nothing

    Private ThumbItems As VisualCollection = Nothing

    Public Sub New(element As UIElement)
        MyBase.New(element)

        Me.ThumbItems = New VisualCollection(Me)

        Me.BuildAdornerCorner(Me.TopThumb, Cursors.SizeNS)
        Me.BuildAdornerCorner(Me.BottomThumb, Cursors.SizeNS)
        Me.BuildAdornerCorner(Me.LeftThumb, Cursors.SizeWE)
        Me.BuildAdornerCorner(Me.RightThumb, Cursors.SizeWE)

        Me.BuildAdornerCorner(Me.TopLeftThumb, Cursors.SizeNWSE)
        Me.BuildAdornerCorner(Me.TopRightThumb, Cursors.SizeNESW)
        Me.BuildAdornerCorner(Me.BottomLeftThumb, Cursors.SizeNESW)
        Me.BuildAdornerCorner(Me.BottomRightThumb, Cursors.SizeNWSE)

        AddHandler Me.TopThumb.DragDelta, AddressOf Me.TopThumb_DragDelta
        AddHandler Me.BottomThumb.DragDelta, AddressOf Me.BottomThumb_DragDelta
        AddHandler Me.LeftThumb.DragDelta, AddressOf Me.LeftThumb_DragDelta
        AddHandler Me.RightThumb.DragDelta, AddressOf Me.RightThumb_DragDelta

        AddHandler Me.TopLeftThumb.DragDelta, AddressOf Me.TopLeftThumb_DragDelta
        AddHandler Me.TopRightThumb.DragDelta, AddressOf Me.TopRightThumb_DragDelta
        AddHandler Me.BottomLeftThumb.DragDelta, AddressOf Me.BottomLeftThumb_DragDelta
        AddHandler Me.BottomRightThumb.DragDelta, AddressOf Me.BottomRightThumb_DragDelta

    End Sub

    Private Sub BuildAdornerCorner(ByRef cornerThumb As Thumb, newCursor As Cursor)

        If cornerThumb IsNot Nothing Then
            Return
        End If

        cornerThumb = New Thumb
        cornerThumb.Cursor = newCursor
        cornerThumb.Width = 10
        cornerThumb.Height = 10
        cornerThumb.Opacity = 0.4
        cornerThumb.Background = New SolidColorBrush(Colors.MediumBlue)
        Me.ThumbItems.Add(cornerThumb)

    End Sub

    Private Sub TopThumb_DragDelta(sender As Object, e As DragDeltaEventArgs)

        Dim element = TryCast(Me.AdornedElement, FrameworkElement)
        Dim cornerThumb = TryCast(e.Source, Thumb)

        If element Is Nothing OrElse cornerThumb Is Nothing Then
            Return
        End If

        Me.EnforceSize(element)

        Dim oldHeight = element.Height
        Dim newHeight = Math.Max(element.Height - e.VerticalChange, cornerThumb.DesiredSize.Height)
        Dim oldTop = Canvas.GetTop(element)

        Canvas.SetTop(element, oldTop - (newHeight - oldHeight))
        element.Height = newHeight

    End Sub

    Private Sub BottomThumb_DragDelta(sender As Object, e As DragDeltaEventArgs)

        Dim element = TryCast(Me.AdornedElement, FrameworkElement)
        Dim cornerThumb = TryCast(e.Source, Thumb)

        If element Is Nothing OrElse cornerThumb Is Nothing Then
            Return
        End If

        Me.EnforceSize(element)

        element.Height = Math.Max(element.Height + e.VerticalChange, cornerThumb.DesiredSize.Height)

    End Sub

    Private Sub LeftThumb_DragDelta(sender As Object, e As DragDeltaEventArgs)

        Dim element = TryCast(Me.AdornedElement, FrameworkElement)
        Dim cornerThumb = TryCast(e.Source, Thumb)

        If element Is Nothing OrElse cornerThumb Is Nothing Then
            Return
        End If

        Me.EnforceSize(element)

        Dim oldWidth = element.Width
        Dim newWidth = Math.Max(element.Width - e.HorizontalChange, cornerThumb.DesiredSize.Width)
        Dim oldLeft = Canvas.GetLeft(element)

        Canvas.SetLeft(element, oldLeft - (newWidth - oldWidth))
        element.Width = newWidth

    End Sub

    Private Sub RightThumb_DragDelta(sender As Object, e As DragDeltaEventArgs)

        Dim element = TryCast(Me.AdornedElement, FrameworkElement)
        Dim cornerThumb = TryCast(e.Source, Thumb)

        If element Is Nothing OrElse cornerThumb Is Nothing Then
            Return
        End If

        Me.EnforceSize(element)

        element.Width = Math.Max(element.Width + e.HorizontalChange, cornerThumb.DesiredSize.Width)

    End Sub

    Private Sub TopLeftThumb_DragDelta(sender As Object, e As DragDeltaEventArgs)

        Dim element = TryCast(Me.AdornedElement, FrameworkElement)
        Dim cornerThumb = TryCast(e.Source, Thumb)

        If element Is Nothing OrElse cornerThumb Is Nothing Then
            Return
        End If

        Me.EnforceSize(element)

        Dim oldWidth = element.Width
        Dim newWidth = Math.Max(element.Width - e.HorizontalChange, cornerThumb.DesiredSize.Width)
        Dim oldLeft = Canvas.GetLeft(element)

        Canvas.SetLeft(element, oldLeft - (newWidth - oldWidth))
        element.Width = newWidth

        Dim oldHeight = element.Height
        Dim newHeight = Math.Max(element.Height - e.VerticalChange, cornerThumb.DesiredSize.Height)
        Dim oldTop = Canvas.GetTop(element)

        Canvas.SetTop(element, oldTop - (newHeight - oldHeight))
        element.Height = newHeight

    End Sub

    Private Sub TopRightThumb_DragDelta(sender As Object, e As DragDeltaEventArgs)

        Dim element = TryCast(Me.AdornedElement, FrameworkElement)
        Dim cornerThumb = TryCast(e.Source, Thumb)

        If element Is Nothing OrElse cornerThumb Is Nothing Then
            Return
        End If

        Me.EnforceSize(element)

        element.Width = Math.Max(element.Width + e.HorizontalChange, cornerThumb.DesiredSize.Width)

        Dim oldHeight = element.Height
        Dim newHeight = Math.Max(element.Height - e.VerticalChange, cornerThumb.DesiredSize.Height)
        Dim oldTop = Canvas.GetTop(element)

        Canvas.SetTop(element, oldTop - (newHeight - oldHeight))
        element.Height = newHeight

    End Sub

    Private Sub BottomLeftThumb_DragDelta(sender As Object, e As DragDeltaEventArgs)

        Dim element = TryCast(Me.AdornedElement, FrameworkElement)
        Dim cornerThumb = TryCast(e.Source, Thumb)

        If element Is Nothing OrElse cornerThumb Is Nothing Then
            Return
        End If

        Me.EnforceSize(element)

        Dim oldWidth = element.Width
        Dim newWidth = Math.Max(element.Width - e.HorizontalChange, cornerThumb.DesiredSize.Width)
        Dim oldLeft = Canvas.GetLeft(element)

        Canvas.SetLeft(element, oldLeft - (newWidth - oldWidth))
        element.Width = newWidth

        element.Height = Math.Max(element.Height + e.VerticalChange, cornerThumb.DesiredSize.Height)

    End Sub

    Private Sub BottomRightThumb_DragDelta(sender As Object, e As DragDeltaEventArgs)

        Dim element = TryCast(Me.AdornedElement, FrameworkElement)
        Dim cornerThumb = TryCast(e.Source, Thumb)

        If element Is Nothing OrElse cornerThumb Is Nothing Then
            Return
        End If

        Me.EnforceSize(element)

        element.Width = Math.Max(element.Width + e.HorizontalChange, cornerThumb.DesiredSize.Width)
        element.Height = Math.Max(element.Height + e.VerticalChange, cornerThumb.DesiredSize.Height)

    End Sub

    Private Sub EnforceSize(element As FrameworkElement)

        If element.Width.Equals(Double.NaN) Then
            element.Width = element.DesiredSize.Width
        End If

        If element.Height.Equals(Double.NaN) Then
            element.Height = element.DesiredSize.Height
        End If

        Dim parentElement = TryCast(element.Parent, FrameworkElement)
        If parentElement Is Nothing Then
            Return
        End If

        element.MaxWidth = parentElement.ActualWidth
        element.MaxHeight = parentElement.ActualHeight

    End Sub

    Protected Overrides Function ArrangeOverride(finalSize As Size) As Size

        Dim desiredWidth = Me.AdornedElement.DesiredSize.Width
        Dim desiredHeight = Me.AdornedElement.DesiredSize.Height
        Dim adornerWidth = Me.DesiredSize.Width
        Dim adornerHeight = Me.DesiredSize.Height

        'Me.TopLeftThumb.Arrange(New Rect(-adornerWidth / 2, -adornerHeight / 2, adornerWidth, adornerHeight))
        'Me.TopRightThumb.Arrange(New Rect(desiredWidth - adornerWidth / 2, -adornerHeight / 2, adornerWidth, adornerHeight))
        'Me.BottomLeftThumb.Arrange(New Rect(-adornerWidth / 2, desiredHeight - adornerHeight / 2, adornerWidth, adornerHeight))
        'Me.BottomRightThumb.Arrange(New Rect(desiredWidth - adornerWidth / 2, desiredHeight - adornerHeight / 2, adornerWidth, adornerHeight))
        Dim newX As Double = 0
        Dim newY As Double = 0

        ' x, y は、図形の中心を(0, 0)として、そこから x 分増減の移動、y 分増減の移動しているっぽい？

        newX = 0
        newY = -(adornerHeight / 2) - (Me.TopThumb.Height / 2)
        Me.TopThumb.Arrange(New Rect(newX, newY, adornerWidth, adornerHeight))

        newX = 0
        newY = desiredHeight - (adornerHeight / 2) + (Me.BottomThumb.Height / 2)
        Me.BottomThumb.Arrange(New Rect(newX, newY, adornerWidth, adornerHeight))

        newX = -(adornerWidth / 2) - (Me.LeftThumb.Width / 2)
        newY = 0
        Me.LeftThumb.Arrange(New Rect(newX, newY, adornerWidth, adornerHeight))

        newX = desiredWidth - (adornerWidth / 2) + (Me.RightThumb.Width / 2)
        newY = 0
        Me.RightThumb.Arrange(New Rect(newX, newY, adornerWidth, adornerHeight))







        newX = -(adornerWidth / 2) - (Me.TopLeftThumb.Width / 2)
        newY = -(adornerHeight / 2) - (Me.TopLeftThumb.Height / 2)
        Me.TopLeftThumb.Arrange(New Rect(newX, newY, adornerWidth, adornerHeight))

        newX = desiredWidth - (adornerWidth / 2) + (Me.TopRightThumb.Width / 2)
        newY = -(adornerHeight / 2) - (Me.TopRightThumb.Height / 2)
        Me.TopRightThumb.Arrange(New Rect(newX, newY, adornerWidth, adornerHeight))

        newX = -(adornerWidth / 2) - (Me.BottomLeftThumb.Width / 2)
        newY = desiredHeight - (adornerHeight / 2) + (Me.BottomLeftThumb.Height / 2)
        Me.BottomLeftThumb.Arrange(New Rect(newX, newY, adornerWidth, adornerHeight))

        newX = desiredWidth - (adornerWidth / 2) + (Me.BottomRightThumb.Width / 2)
        newY = desiredHeight - (adornerHeight / 2) + (Me.BottomRightThumb.Height / 2)
        Me.BottomRightThumb.Arrange(New Rect(newX, newY, adornerWidth, adornerHeight))

        Return finalSize

    End Function

    Protected Overrides ReadOnly Property VisualChildrenCount As Integer
        Get
            Return Me.ThumbItems.Count
        End Get
    End Property

    Protected Overrides Function GetVisualChild(index As Integer) As Visual
        Return Me.ThumbItems(index)
    End Function

End Class
