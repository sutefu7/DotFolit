Imports System.Windows.Controls.Primitives
Imports DotFolit.Petzold.Media2D


' このコントロールは、Canvas コントロール上に配置して使われることを前提に作成しています。
' それ以外のコンテナコントロール（Grid 他）上に配置した場合の考慮は、含まれていません。


Public Class ResizableThumb
    Inherits Thumb

#Region "フィールド、プロパティ"

    Private AdornmentLayer As AdornerLayer = Nothing

    Public Property StartLines As List(Of ArrowLine) = Nothing
    Public Property EndLines As List(Of ArrowLine) = Nothing

#End Region

#Region "UseAdorner 依存関係プロパティ"

    Public Shared ReadOnly UseAdornerProperty As DependencyProperty =
        DependencyProperty.Register(
        "UseAdorner",
        GetType(Boolean),
        GetType(ResizableThumb),
        New FrameworkPropertyMetadata(False, AddressOf UseAdornerPropertyChanged))

    Public Property UseAdorner As Boolean
        Get
            Return CType(Me.GetValue(UseAdornerProperty), Boolean)
        End Get
        Set(value As Boolean)
            Me.SetValue(UseAdornerProperty, value)
        End Set
    End Property

    Private Shared Sub UseAdornerPropertyChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)

        Dim changedThumb = TryCast(d, ResizableThumb)
        If changedThumb Is Nothing Then
            Return
        End If

        ' UseAdorner プロパティが False なのに、IsSelected プロパティが True の場合、False に戻す
        Dim useAdornerValue = CType(e.NewValue, Boolean)
        If Not useAdornerValue Then

            If changedThumb.IsSelected Then
                changedThumb.IsSelected = False
            End If
            Return

        End If

    End Sub

#End Region

#Region "IsSelected 依存関係プロパティ"

    Public Shared ReadOnly IsSelectedProperty As DependencyProperty =
        DependencyProperty.Register(
        "IsSelected",
        GetType(Boolean),
        GetType(ResizableThumb),
        New FrameworkPropertyMetadata(False, AddressOf IsSelectedPropertyChanged, AddressOf IsSelectedPropertyCoerceValue))

    Public Property IsSelected As Boolean
        Get
            Return CType(Me.GetValue(IsSelectedProperty), Boolean)
        End Get
        Set(value As Boolean)
            Me.SetValue(IsSelectedProperty, value)
        End Set
    End Property

    Private Shared Function IsSelectedPropertyCoerceValue(d As DependencyObject, obj As Object) As Object

        Dim changedThumb = TryCast(d, ResizableThumb)
        If changedThumb Is Nothing Then
            Return obj
        End If

        ' UseAdorner プロパティが False の場合、IsSelected プロパティに True/False をセットしても False として受け取る
        If Not changedThumb.UseAdorner Then
            Return False
        End If

        Return obj

    End Function

    Private Shared Sub IsSelectedPropertyChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)

        Dim changedThumb = TryCast(d, ResizableThumb)
        If changedThumb Is Nothing Then
            Return
        End If

        ' 変更後の値を見て、Adorner 装飾を切り替える
        Dim isSelectedValue = CType(e.NewValue, Boolean)
        If isSelectedValue Then
            ' Adorner を付ける
            changedThumb.AttachAdorner()
        Else
            ' Adorner を外す
            changedThumb.DetachAdorner()
        End If

    End Sub

#End Region

#Region "コンストラクタ"

    Public Sub New()
        MyBase.New()

        Me.StartLines = New List(Of ArrowLine)
        Me.EndLines = New List(Of ArrowLine)

        AddHandler Me.Loaded, AddressOf Me.Thumb_Loaded
        AddHandler Me.DragDelta, AddressOf Me.Thumb_DragDelta

    End Sub

#End Region

#Region "ロード"

    ' コンストラクタで実行すると、xaml例外エラーが発生するため、ロードイベントで実行するように、実行タイミングを変更
    Private Sub Thumb_Loaded(sender As Object, e As RoutedEventArgs)

        ' 親インスタンスを取得
        ' 背景色が未セットの場合、透明色？を塗る（何かを塗っていないと、クリック判定されない現象の対応）
        ' 
        ' WPF:CanvasなどのコントロールはBackground/Fillを明示的に指定しないとマウスイベントが発生しない
        ' https://qiita.com/nossey/items/3cf152c5fc2a2f24f585
        ' 
        Dim parentCanvas = TryCast(Me.Parent, Canvas)
        If parentCanvas.Background Is Nothing Then
            parentCanvas.Background = Brushes.Transparent
        End If

        AddHandler parentCanvas.PreviewMouseLeftButtonDown, AddressOf parentCanvas_PreviewMouseLeftButtonDown


    End Sub

#End Region

#Region "ドラッグ移動"

    Private Sub Thumb_DragDelta(sender As Object, e As DragDeltaEventArgs)

        Dim moveThumb = TryCast(e.Source, ResizableThumb)
        If moveThumb Is Nothing Then
            Return
        End If

        Dim movedLeft = Canvas.GetLeft(moveThumb) + e.HorizontalChange
        Dim movedTop = Canvas.GetTop(moveThumb) + e.VerticalChange

        Canvas.SetLeft(moveThumb, movedLeft)
        Canvas.SetTop(moveThumb, movedTop)

        Me.UpdateLineLocation(moveThumb)

    End Sub

#End Region

#Region "親キャンパス、プレビューマウス、左ボタン、ダウン"

    Private Sub parentCanvas_PreviewMouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs)

        Dim clickedElement = TryCast(e.Source, UIElement)
        If clickedElement Is Nothing Then
            Return
        End If

        If clickedElement Is Me Then
            Me.IsSelected = True
        Else
            Me.IsSelected = False
        End If

    End Sub

#End Region

#Region "メソッド"

    ' AttachAdorner メソッドは直接呼び出さず、IsSelected プロパティを利用してください。
    Private Sub AttachAdorner()

        Me.AdornmentLayer = AdornerLayer.GetAdornerLayer(Me)
        Me.AdornmentLayer.Add(New ResizingAdorner(Me))

    End Sub

    ' DetachAdorner メソッドは直接呼び出さず、IsSelected プロパティを利用してください。
    Private Sub DetachAdorner()

        If Me.AdornmentLayer Is Nothing Then
            Return
        End If

        Dim items = Me.AdornmentLayer.GetAdorners(Me)
        If (items IsNot Nothing) AndAlso (items.Count() <> 0) Then
            Me.AdornmentLayer.Remove(items(0))
        End If

    End Sub

    ' EditorUserControl.xaml.vb, MethodWindow.xaml.vb 側にも同じメソッドがあるので同期すること
    Private Sub UpdateLineLocation(target As ResizableThumb)

        Dim newX = Canvas.GetLeft(target)
        Dim newY = Canvas.GetTop(target)

        Dim newWidth = target.ActualWidth
        Dim newHeight = target.ActualHeight

        If newWidth = 0 Then newWidth = target.DesiredSize.Width
        If newHeight = 0 Then newHeight = target.DesiredSize.Height

        For i As Integer = 0 To target.StartLines.Count - 1
            target.StartLines(i).X1 = newX + newWidth
            target.StartLines(i).Y1 = newY + (newHeight / 2)
        Next

        For i As Integer = 0 To target.EndLines.Count - 1
            target.EndLines(i).X2 = newX
            target.EndLines(i).Y2 = newY + (newHeight / 2)
        Next

    End Sub

#End Region


End Class
