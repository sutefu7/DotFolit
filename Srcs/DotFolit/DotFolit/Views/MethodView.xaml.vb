Imports System.Data
Imports System.IO
Imports ICSharpCode.AvalonEdit
Imports ICSharpCode.AvalonEdit.Folding
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.MSBuild
Imports Microsoft.CodeAnalysis.FindSymbols
Imports DotFolit.Petzold.Media2D

Public Class MethodView

#Region "フィールド、プロパティ"

    Private SelectedThumb As ResizableThumb = Nothing
    Private NextSourceFile As String = String.Empty
    Private NextStartLength As Integer = -1

#End Region

#Region "画面のロード"

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)

        Dim vm = TryCast(Me.DataContext, MethodViewModel)
        Me.AddNew(vm.SelectedNode.FileName, 0)

    End Sub

#End Region

#Region "エディタのキャレット位置移動"

    Private Async Sub Caret_PositionChanged(sender As Object, e As EventArgs)

        Dim texteditor1 = TryCast(sender, TextEditor)
        Dim sourceFile = texteditor1.Document.FileName
        Dim offset = texteditor1.TextArea.Caret.Offset
        Dim menu1 = TryCast(texteditor1.ContextMenu.Items(0), MenuItem)

        Dim sourceTree = MemoryDB.Instance.SyntaxTreeItems.FirstOrDefault(Function(x) x.FilePath = sourceFile)
        Dim si As ISymbol = Nothing

        Try

            For Each compilationItem In MemoryDB.Instance.CompilationItems

                Dim model = compilationItem.GetSemanticModel(sourceTree)
                si = Await SymbolFinder.FindSymbolAtPositionAsync(model, offset, MSBuildWorkspace.Create())

                If si IsNot Nothing AndAlso 0 < si.Locations.Count AndAlso si.Locations(0).IsInSource Then
                    Exit For
                End If

            Next



        Catch ex As AggregateException

            Dim idx = 0
            For Each inner In ex.Flatten().InnerExceptions

                idx += 1
                Debug.WriteLine($"{idx} つ目 --------------------------------------")
                Debug.WriteLine(inner.ToString())
                Debug.WriteLine($"------------------------------------------------")
                Debug.WriteLine("")

            Next


        Catch ex As Exception
            Debug.WriteLine(ex.ToString())
        End Try

        ' 見つからなかった
        If si Is Nothing Then

            Me.SelectedThumb = Nothing
            Me.NextSourceFile = String.Empty
            Me.NextStartLength = -1

            'texteditor1.ContextMenu.Visibility = Visibility.Collapsed
            menu1.IsEnabled = False

            Return

        End If

        ' メソッド以外は対象外ではじく
        If si.Kind = SymbolKind.Method Then

            Me.SelectedThumb = TryCast(TryCast(TryCast(texteditor1.Parent, DockPanel).Parent, Border).TemplatedParent, ResizableThumb)
            Me.NextSourceFile = si.Locations(0).SourceTree?.FilePath
            Me.NextStartLength = si.Locations(0).SourceSpan.Start

            If String.IsNullOrWhiteSpace(Me.NextSourceFile) OrElse Not File.Exists(Me.NextSourceFile) Then
                Return
            End If

            ' メソッド対象以外を右クリックした後（このタイミングでは、コンテキストメニュー非表示で OK なのだが）、
            ' メソッド対象を左クリックするだけで、コンテキストメニューが表示されてしまう現象の対応
            'texteditor1.ContextMenu.Visibility = Visibility.Visible
            menu1.IsEnabled = True

        Else

            Me.SelectedThumb = Nothing
            Me.NextSourceFile = String.Empty
            Me.NextStartLength = -1

            'texteditor1.ContextMenu.Visibility = Visibility.Collapsed
            menu1.IsEnabled = False

        End If

    End Sub

#End Region

#Region "エディタのマウスホイール変更（ズームイン、アウト）"

    ' [WPF]コントロールキーあるいはシフトキーが押されているかどうかを取得する
    ' https://gist.github.com/pinzolo/3080481


    Private Sub texteditor1_PreviewMouseWheel(sender As Object, e As MouseWheelEventArgs)

        Dim isDownLeftControlKey = ((Keyboard.GetKeyStates(Key.LeftCtrl) And KeyStates.Down) = KeyStates.Down)
        Dim isDownRightControlKey = ((Keyboard.GetKeyStates(Key.RightCtrl) And KeyStates.Down) = KeyStates.Down)
        Dim isDownControlKey = isDownLeftControlKey OrElse isDownRightControlKey

        If isDownControlKey Then

            Dim editor = TryCast(sender, TextEditor)
            If 0 < e.Delta Then
                editor.FontSize *= 1.1
            Else
                editor.FontSize /= 1.1
            End If

        End If

    End Sub

#End Region

#Region "このメソッドを表示メニューのクリック"

    Private Sub MenuItem_Click(sender As Object, e As RoutedEventArgs)

        Me.AddNew(Me.NextSourceFile, Me.NextStartLength)

    End Sub

#End Region

#Region "メソッド"

    Private Sub AddNew(sourceFile As String, startLength As Integer)

        ' 表示図形の作成
        Dim newThumb = New ResizableThumb
        newThumb.UseAdorner = True
        newThumb.Template = TryCast(Me.Resources("EditorTemplate"), ControlTemplate)
        newThumb.ApplyTemplate()
        newThumb.UpdateLayout()

        ' ソース全体が長すぎると見づらい、探しづらい、理解しづらい
        ' 仮対応として固定サイズで表示する。見づらかったらリサイズしてもらう運用の方が、まだマシと判断
        newThumb.Width = 640
        newThumb.Height = 480

        ' タイトルをセット
        Dim textblock1 = TryCast(newThumb.Template.FindName("textblock1", newThumb), TextBlock)
        'textblock1.Text = $"{Path.GetDirectoryName(sourceFile)}/{Path.GetFileName(sourceFile)}"

        Dim fi = New FileInfo(sourceFile)
        textblock1.Text = $"{fi.Directory.Name}/{fi.Name}"

        ' 調査用
        'textblock1.Text = $"{sourceFile}"      

        ' ソースをセット
        Dim texteditor1 = TryCast(newThumb.Template.FindName("texteditor1", newThumb), TextEditor)
        texteditor1.Text = File.ReadAllText(sourceFile, EncodeResolver.GetEncoding(sourceFile))
        texteditor1.Document.FileName = sourceFile

        ' キャレット位置を、メンバー定義位置へ移動
        texteditor1.TextArea.Caret.Offset = startLength

        ' メンバー定義位置が見えるようにスクロール
        ' （いまいちうまく扱えないスクロール処理
        ' TextEditor.ScrollToVerticalOffset メソッドと、TextEditor.ScrollToLine メソッドの組み合わせで、うまくスクロールしてくれた）
        texteditor1.ScrollToVerticalOffset(startLength)

        ' メソッド定義の開始行が見やすいように、２行分上に表示する
        Dim jumpLine = texteditor1.Document.GetLineByOffset(startLength).LineNumber
        If 0 <= jumpLine - 2 Then
            jumpLine -= 2
        End If
        texteditor1.ScrollToLine(jumpLine)

        ' XAML 上で設定していない部分の設定
        ' タブはスペース変換して表示する
        texteditor1.Options.ConvertTabsToSpaces = True

        ' 現在行の背景色を表示する
        texteditor1.Options.HighlightCurrentLine = True

        ' ソースを表示してから、折りたたみ機能を設定
        Dim strategy = New VBNetFoldingStrategy
        Dim manager = FoldingManager.Install(texteditor1.TextArea)
        strategy.UpdateFoldings(manager, texteditor1.Document)

        ' 表示位置をセット
        Dim pos = Me.GetNewLocation()
        Canvas.SetLeft(newThumb, pos.X)
        Canvas.SetTop(newThumb, pos.Y)

        ' マウスホイールイベント、キャレット移動イベントの購読
        AddHandler texteditor1.PreviewMouseWheel, AddressOf Me.texteditor1_PreviewMouseWheel
        'AddHandler texteditor1.TextArea.Caret.PositionChanged, AddressOf Me.Caret_PositionChanged
        AddHandler texteditor1.TextArea.Caret.PositionChanged, Sub(sender, e)

                                                                   ' 通常のままだと Caret コントロールが渡されてくるのだが、ここから texteditor1 コントロールへさかのぼることが出来ない
                                                                   ' texteditor1 コントロールを取得したいので、イベントハンドラをトラップして、渡してしまう
                                                                   sender = texteditor1
                                                                   Me.Caret_PositionChanged(sender, e)

                                                               End Sub

        Dim menu1 = TryCast(newThumb.Template.FindName("AddNewMenu", newThumb), MenuItem)
        AddHandler menu1.Click, AddressOf Me.MenuItem_Click

        ' キャンバスに登録
        Me.canvas1.Children.Add(newThumb)

        ' コネクタの接続
        ' 矢印線でつながれていると、呼び出し元、呼び出し先が分かりやすいのだが、不要か？
        If Me.SelectedThumb Is Nothing Then
            Return
        End If

        ' 前回と今回の図形同士を、矢印線でつなげる
        Dim arrow = New ArrowLine
        arrow.Stroke = Brushes.Green
        arrow.StrokeThickness = 1
        Me.canvas1.Children.Add(arrow)

        Me.SelectedThumb.StartLines.Add(arrow)
        newThumb.EndLines.Add(arrow)

        Me.UpdateLineLocation(Me.SelectedThumb)
        Me.UpdateLineLocation(newThumb)

        ' なぜか、最後の図形だけ、矢印線が左上の角を指してしまう不具合
        ' → ActualWidth, ActualHeight が 0 だから。いったん画面表示させないとダメか？
        ' → Measure メソッドを呼び出して、希望サイズを更新する。こちらで矢印線の位置を調整する
        Dim newSize = New Size(Me.canvas1.ActualWidth, Me.canvas1.ActualHeight)
        Me.canvas1.Measure(newSize)
        Me.UpdateLineLocation(newThumb)

    End Sub

    Private Function GetNewLocation() As Point

        ' 最も右下に位置する ResizableThumb を探す
        Dim items = Me.canvas1.Children.OfType(Of ResizableThumb)()
        If items.Count() = 0 Then
            Return New Point(10, 10)
        End If

        ' 既に表示されている図形のうち、１つ目の図形位置を基準として、今回の図形位置を計算する
        Dim item = If(Me.SelectedThumb Is Nothing, items(0), Me.SelectedThumb)
        Dim newWidth As Double = item.ActualWidth
        Dim newHeight As Double = item.ActualHeight

        Dim newX As Double = Canvas.GetLeft(item) + newWidth + 40
        Dim newY As Double = Canvas.GetTop(item)

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

                Dim currentWidth = item.ActualWidth
                Dim currentHeight = item.ActualHeight

                Dim currentRect = New Rect(currentX, currentY, currentWidth, currentHeight)
                Dim newRect = New Rect(newX, newY, newWidth, newHeight)

                Select Case True
                    Case currentRect.Contains(newRect.TopLeft),
                     currentRect.Contains(newRect.TopRight),
                     currentRect.Contains(newRect.BottomLeft),
                     currentRect.Contains(newRect.BottomRight)

                        ' 重なっている図形の【下+隙間】まで移動
                        found = True
                        newY = currentY + currentHeight + 10

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
