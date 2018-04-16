Imports System.Collections.ObjectModel
Imports DotFolit.Petzold.Media2D

Public Class BindableCanvas
    Inherits Canvas

#Region "ItemsSource 依存関係プロパティ"

    Public Shared ReadOnly ItemsSourceProperty As DependencyProperty =
        DependencyProperty.Register(
        "ItemsSource",
        GetType(IEnumerable),
        GetType(BindableCanvas),
        New PropertyMetadata(AddressOf ItemsSourcePropertyChanged))

    Private Shared Sub ItemsSourcePropertyChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)

        Dim ctrl = TryCast(d, BindableCanvas)
        If 0 < ctrl.Children.Count Then
            'ctrl.Children.Clear()
            ' 一度、クラスの継承関係図を表示させた後は、キャレット移動しても更新しないように対応
            ' ただし、これにより BindableCanvas の機能として、汎用性が無くなった（元々 ViewModel を ObservableCollection(Of InheritsItemModel) 専用で扱っているから開発当初からか）
            Return
        End If

        ctrl.ItemsSource = TryCast(e.NewValue, IEnumerable)
        ctrl.AddChildren()

    End Sub

    Public Property ItemsSource As IEnumerable
        Get
            Return TryCast(GetValue(ItemsSourceProperty), IEnumerable)
        End Get
        Set(value As IEnumerable)
            SetValue(ItemsSourceProperty, value)
        End Set
    End Property

    ' クラスの継承関係図は、最終的にはコードビハインド頼り
    ' VM に対応する V を作成して、バインド済みの V のリストを、Children にセット
    Private Sub AddChildren()

        ' ItemsSource = List<ViewModel>
        ' VM -> View の生成とバインドをする
        ' List<View> を Children にセットする、矢印もセットする
        Dim vmItems = TryCast(Me.ItemsSource, ObservableCollection(Of InheritsItemModel))
        If vmItems.Count = 0 Then
            Return
        End If

        For Each vmItem In vmItems
            Me.AddChildren(vmItem)
        Next

    End Sub

    Private Sub AddChildren(vm As InheritsItemModel)

        Dim thumbTemplate = TryCast(Me.Resources("ClassMemberTemplate"), ControlTemplate)
        Dim newThumb = New ResizableThumb
        newThumb.Template = thumbTemplate
        newThumb.DataContext = vm

        Dim newSize = New Size(50000, 50000)
        Me.Measure(newSize)

        newThumb.ApplyTemplate()
        newThumb.UpdateLayout()

        ' Canvas 上の表示位置の設定
        Dim pos = Me.GetNewLocation()
        Canvas.SetLeft(newThumb, pos.X)
        Canvas.SetTop(newThumb, pos.Y)

        ' キャンバスに登録する前に、１つ前に登録した図形を取得して置く
        Dim previousThumb As ResizableThumb = Nothing
        If Me.Children.OfType(Of ResizableThumb)().Count <> 0 Then

            Dim items = Me.Children.OfType(Of ResizableThumb)()
            previousThumb = items.LastOrDefault()

        End If

        ' 今回分の図形を新規登録
        Me.Children.Add(newThumb)

        If previousThumb Is Nothing Then
            Return
        End If

        ' 前回と今回の図形同士を、矢印線でつなげる
        Dim arrow = New ArrowLine
        arrow.Stroke = Brushes.Green
        arrow.StrokeThickness = 1
        Me.Children.Add(arrow)

        previousThumb.StartLines.Add(arrow)
        newThumb.EndLines.Add(arrow)

        Me.UpdateLineLocation(previousThumb)
        Me.UpdateLineLocation(newThumb)

        ' なぜか、最後の図形だけ、矢印線が左上の角を指してしまう不具合
        ' → ActualWidth, ActualHeight が 0 だから。いったん画面表示させないとダメか？
        ' → Measure メソッドを呼び出して、希望サイズを更新する。こちらで矢印線の位置を調整する
        newSize = New Size(50000, 50000)
        Me.Measure(newSize)
        Me.UpdateLineLocation(newThumb)

    End Sub

    Private Function GetNewLocation() As Point

        ' ActualWidth, ActualHeight / DesiredSize.Width, DesiredSize.Height
        ' いまいち 0 から値が更新されるタイミングが分からない

        ' 最も右下に位置する ResizableThumb を探す
        Dim items = Me.Children.OfType(Of ResizableThumb)()
        If items.Count() = 0 Then
            Return New Point(10, 10)
        End If

        ' 既に表示されている図形のうち、１つ目の図形位置を基準として、今回の図形位置を計算する
        Dim item = items(0)
        Dim newWidth As Double = item.DesiredSize.Width
        Dim newHeight As Double = item.DesiredSize.Height

        Dim newX As Double = Canvas.GetLeft(item) + newWidth + 40
        Dim newY As Double = Canvas.GetTop(item) + 40 ' Math.Min(40, newHeight / 2)

        ' 予測計算した位置・コントロールの大きさの内に、すでに他のコントロールが配置されていないかチェック
        ' 一部が重なり合う場合、重ならないように予測位置を修正して、再度全チェック
        Dim found = False

        While True

            ' 初期化してチャレンジ、または再チャレンジ
            found = False
            For i As Integer = 1 To items.Count - 1

                item = items(i)
                Dim currentX = Canvas.GetLeft(item)
                Dim currentY = Canvas.GetTop(item)

                Dim currentWidth = item.DesiredSize.Width
                Dim currentHeight = item.DesiredSize.Height

                Dim currentRect = New Rect(currentX, currentY, currentWidth, currentHeight)
                Dim newRect = New Rect(newX, newY, newWidth, newHeight)

                Select Case True
                    Case currentRect.Contains(newRect.TopLeft),
                     currentRect.Contains(newRect.TopRight),
                     currentRect.Contains(newRect.BottomLeft),
                     currentRect.Contains(newRect.BottomRight)

                        ' 重なっている図形の【右下ちょい上くらい】まで移動
                        found = True
                        newX = currentX + currentWidth + 40
                        newY = currentY + 40 ' Math.Min(40, currentHeight / 2)

                End Select

            Next

            ' 既存配置しているコントロールリスト全てに重ならなかったので、この位置で決定
            If Not found Then
                Exit While
            End If

            ' １つ以上重なっていたことにより、予測位置を修正したので、もう一度コントロールリスト全部と再チェック
        End While

        Dim result = New Point(newX, newY)
        Return result

    End Function

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
