Imports System.Data
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Reflection
Imports System.Windows.Controls.Primitives

Imports ICSharpCode.AvalonEdit
Imports ICSharpCode.AvalonEdit.Folding

Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.MSBuild
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax
Imports Microsoft.CodeAnalysis.FindSymbols

Imports DotFolit.Petzold.Media2D


Public Class MethodWindow

#Region "フィールド、プロパティ"

    Public Property StartSourceFile As String = String.Empty

#End Region

#Region "画面のロード"

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)

        Me.AddNew(Me.StartSourceFile, 0)

    End Sub

#End Region

#Region "Thumb コントロールのドラッグアンドドロップ移動イベント"

    ' なんで、あらぶるんだろう？

    Private Sub Thumb_DragDelta(sender As Object, e As DragDeltaEventArgs)

        Dim moveThumb = TryCast(e.Source, ResizableThumb)
        If moveThumb Is Nothing Then
            Return
        End If

        Canvas.SetLeft(moveThumb, Canvas.GetLeft(moveThumb) + e.HorizontalChange)
        Canvas.SetTop(moveThumb, Canvas.GetTop(moveThumb) + e.VerticalChange)

        Me.UpdateLineLocation(moveThumb)

    End Sub

#End Region

#Region "Menu の隣に追加クリックイベント"

    Private Sub Menu_Click_AddNew(sender As Object, e As RoutedEventArgs)

        'Me.AddNew()

    End Sub

#End Region

#Region "Menu の最前面に移動クリックイベント"

    Private Sub Menu_Click_ChangeZIndexToMostTop(sender As Object, e As RoutedEventArgs)

        'Canvas.SetZIndex(Me.SelectedThumb, 1)

    End Sub

#End Region

#Region "エディタのキャレット位置移動"

    Private Async Sub Caret_PositionChanged(sender As Object, e As EventArgs)

        Dim texteditor1 = TryCast(sender, TextEditor)
        Dim thumb1 = TryCast(TryCast(TryCast(texteditor1.Parent, DockPanel).Parent, Border).TemplatedParent, ResizableThumb)
        Dim sourceFile = texteditor1.Document.FileName
        Dim offset = texteditor1.TextArea.Caret.Offset

        Dim sourceTree = MemoryDB.Instance.SyntaxTreeItems.FirstOrDefault(Function(x) x.FilePath = sourceFile)
        Dim model = MemoryDB.Instance.CompilationItem.GetSemanticModel(sourceTree)
        Dim si As ISymbol = Nothing

        Try

            si = Await SymbolFinder.FindSymbolAtPositionAsync(model, offset, MSBuildWorkspace.Create())

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

        If si Is Nothing Then
            Debug.WriteLine("not found")
            Return
        End If

        ' メソッド以外は対象外ではじく
        If si.Kind <> SymbolKind.Method Then
            Return
        End If


        'If si.Locations(0).Kind = LocationKind.SourceFile Then
        '    ' 内部ソースにある場合？

        '    Dim callerFile = si.Locations(0).SourceTree.FilePath
        '    Dim startLength = si.Locations(0).SourceSpan.Start
        '    Console.WriteLine($"{callerFile}({startLength})")

        '    Me.AddNew(callerFile, startLength)

        'Else
        '    ' 外部dll側にある場合？
        '    Dim target = si.ToString()
        '    Dim result = Me.TreeItems.FirstOrDefault(Function(x) x.GetRoot().DescendantNodes().Any(Function(y) y.ToString() = target))
        '    Dim result2 = result.GetRoot().DescendantNodes().FirstOrDefault(Function(y) y.ToString() = target)

        '    Dim callerFile = result2.SyntaxTree.FilePath
        '    Dim startLength = result2.Span.Start
        '    Console.WriteLine($"{callerFile}({startLength})")

        '    Me.AddNew(callerFile, startLength)

        'End If



        ' 通常はソースファイル＝vbproj ファイルがあるフォルダなのだが、
        ' プロジェクト内にサブフォルダを作成している場合、さかのぼってフォルダパスを取得する
        ' 「vbproj ファイルがあるフォルダ」をもとに、同じ名前空間かどうかを判断したいため
        Dim sourceDir = Path.GetDirectoryName(sourceFile)
        While True

            If Directory.EnumerateFiles(sourceDir, "*.vbproj").Any() Then
                Exit While
            End If
            sourceDir = Path.GetDirectoryName(sourceDir)

        End While

        Dim signature = si.ToString()
        Dim candidateTrees = MemoryDB.Instance.SyntaxTreeItems.Where(Function(x) x.GetRoot().DescendantNodes().Any(Function(y) y.ToString() = signature))
        Dim foundSignature = False

        If candidateTrees.Count() = 1 Then

            Dim node = candidateTrees(0).GetRoot().DescendantNodes().FirstOrDefault(Function(x) x.ToString() = signature)
            Dim callerFile = node.SyntaxTree.FilePath
            Dim startLength = node.Span.Start
            Me.AddNew(callerFile, startLength, thumb1)
            foundSignature = True
        End If

        ' 可能性が複数ある場合、以下の優先度で探すのはどうか？
        ' １．同じフォルダ内（プロジェクト内）、または子フォルダ内にファイル名がある
        ' ２．違うフォルダ内（同ソリューションフォルダ内ではあるが、別フォルダ内）にファイル名がある

        ' 動作結果を見ると、ISymbol.ContainingAssembly がMyCompilation（作成時に名付けた名称）の場合、同プロジェクト内にありそう、
        ' それ以外の場合、別プロジェクト内にありそう、ということみたい？

        ' １．同プロジェクト内
        If Not foundSignature Then

            For Each candidateTree In candidateTrees

                If si.ContainingAssembly.ToString().Contains("MyCompilation") Then

                    If candidateTree.FilePath.StartsWith(sourceDir) Then

                        Dim node = candidateTree.GetRoot().DescendantNodes().FirstOrDefault(Function(x) x.ToString() = signature)
                        Dim callerFile = node.SyntaxTree.FilePath
                        Dim startLength = node.Span.Start
                        Me.AddNew(callerFile, startLength, thumb1)

                        foundSignature = True
                        Exit For

                    End If

                End If

            Next

        End If

        ' ２．別プロジェクト内
        If Not foundSignature Then

            For Each candidateTree In candidateTrees

                If Not candidateTree.FilePath.StartsWith(sourceDir) Then

                    Dim node = candidateTree.GetRoot().DescendantNodes().FirstOrDefault(Function(x) x.ToString() = signature)
                    Dim callerFile = node.SyntaxTree.FilePath
                    Dim startLength = node.Span.Start
                    Me.AddNew(callerFile, startLength, thumb1)

                    foundSignature = True
                    Exit For

                End If

            Next

        End If

        If Not foundSignature Then
            Return
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

#Region "メソッド"

    Private Sub AddNew(sourceFile As String, startLength As Integer, Optional selectedThumb As ResizableThumb = Nothing)

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
        textblock1.Text = $"{Path.GetFileName(sourceFile)}"

        ' 調査用
        'textblock1.Text = $"{sourceFile}"      

        ' ソースをセット
        Dim texteditor1 = TryCast(newThumb.Template.FindName("texteditor1", newThumb), TextEditor)
        texteditor1.Document.Text = File.ReadAllText(sourceFile, EncodeResolver.GetEncoding(sourceFile))
        texteditor1.Document.FileName = sourceFile

        ' キャレット位置を、メンバー定義位置へ移動
        texteditor1.TextArea.Caret.Offset = startLength

        ' メンバー定義位置が見えるようにスクロール
        ' （いまいちうまく扱えないスクロール処理
        ' EditorUserControl.xaml.vb/treeview1_SelectedItemChanged メソッドに書いたやり方と同じだと、うまくいかない
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
        Dim pos = Me.GetNewLocation(selectedThumb)
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

        ' 移動イベントの購読
        AddHandler newThumb.DragDelta, AddressOf Me.Thumb_DragDelta

        ' キャンバスに登録
        Me.MethodCanvas.Children.Add(newThumb)

        ' コネクタの接続
        ' 矢印線でつながれていると、呼び出し元、呼び出し先が分かりやすいのだが、不要か？
        If selectedThumb Is Nothing Then
            Return
        End If

        ' 前回と今回の図形同士を、矢印線でつなげる
        Dim arrow = New ArrowLine
        arrow.Stroke = Brushes.Green
        arrow.StrokeThickness = 1
        MethodCanvas.Children.Add(arrow)

        selectedThumb.StartLines.Add(arrow)
        newThumb.EndLines.Add(arrow)

        Me.UpdateLineLocation(selectedThumb)
        Me.UpdateLineLocation(newThumb)

        ' なぜか、最後の図形だけ、矢印線が左上の角を指してしまう不具合
        ' → ActualWidth, ActualHeight が 0 だから。いったん画面表示させないとダメか？
        ' → Measure メソッドを呼び出して、希望サイズを更新する。こちらで矢印線の位置を調整する
        Dim newSize = New Size(MethodCanvas.ActualWidth, MethodCanvas.ActualHeight)
        MethodCanvas.Measure(newSize)
        Me.UpdateLineLocation(newThumb)

    End Sub

    Private Function GetNewLocation(Optional selectedThumb As ResizableThumb = Nothing) As Point

        ' 最も右下に位置する ResizableThumb を探す
        Dim items = Me.MethodCanvas.Children.OfType(Of ResizableThumb)()
        If items.Count() = 0 Then
            Return New Point(10, 10)
        End If

        ' 既に表示されている図形のうち、１つ目の図形位置を基準として、今回の図形位置を計算する
        Dim item = If(selectedThumb Is Nothing, items(0), selectedThumb)
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

    ' ResizableThumb.vb, EditorUserControl.xaml.vb 側にも同じメソッドがあるので同期すること
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
