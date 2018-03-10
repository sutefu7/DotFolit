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


Public Class MethodWindow

#Region "フィールド、プロパティ"

    Public Property SolutionModel As TreeViewItemModel = Nothing
    Public Property SourceModel As TreeViewItemModel = Nothing

    Private CurrentCompilation As VisualBasicCompilation = Nothing
    Private TreeItems As List(Of SyntaxTree) = Nothing

    Private CallerSourceFile As String = String.Empty
    Private CallerStartLength As Integer = -1

#End Region

#Region "画面のロード"

    Private Async Sub Window_Loaded(sender As Object, e As RoutedEventArgs)

        Await Me.InitializeModel()

    End Sub

    Private Async Function InitializeModel() As Task

        ' 本来は MainWindow.xaml.vb 側の D&D 時に１回だけでやるべきだったかも
        ' 画面表示するたびに全読み込みするのは富豪プログラミングすぎたかも

        Dim task1 = Task.Run(Sub()

                                 Dim dllItems = New List(Of MetadataReference)
                                 Dim srcItems = New List(Of SyntaxTree)

                                 Try

                                     Dim solutionFile = Me.SolutionModel.FileName
                                     Dim msWorkspace = MSBuildWorkspace.Create()
                                     Dim solution = msWorkspace.OpenSolutionAsync(solutionFile).Result

                                     ' プロジェクト数分
                                     For Each project In solution.Projects

                                         ' 参照dll
                                         For Each metaRef In project.MetadataReferences

                                             If Not dllItems.Any(Function(x) x.Display = metaRef.Display) Then
                                                 dllItems.Add(MetadataReference.CreateFromFile(metaRef.Display))
                                             End If

                                         Next

                                         ' ソース
                                         For Each document In project.Documents

                                             Dim source = File.ReadAllText(document.FilePath, EncodeResolver.GetEncoding(document.FilePath))
                                             Dim tree = VisualBasicSyntaxTree.ParseText(source,, document.FilePath)
                                             srcItems.Add(tree)

                                         Next

                                         ' アセンブリファイル
                                         If Not dllItems.Any(Function(x) x.Display = project.OutputFilePath) Then
                                             dllItems.Add(MetadataReference.CreateFromFile(project.OutputFilePath))
                                         End If

                                     Next


                                 Catch ex As ReflectionTypeLoadException

                                     Console.WriteLine(ex.ToString())

                                     For Each inner In ex.LoaderExceptions
                                         Console.WriteLine(inner.ToString())
                                     Next

                                 Catch ex As Exception
                                     Console.WriteLine(ex.ToString())
                                 End Try

                                 Dim compilation = VisualBasicCompilation.Create("MyCompilation", srcItems, dllItems)

                                 ' クラス内からアクセス出来るようにセット
                                 Me.CurrentCompilation = compilation
                                 Me.TreeItems = srcItems

                                 Me.Dispatcher.BeginInvoke(Sub()

                                                               Dim sourceFile = SourceModel.FileName
                                                               Me.AddNew(sourceFile, 0)

                                                           End Sub)

                             End Sub)


        Dim dlg = New ProgressWindow
        dlg.Owner = Me
        dlg.Topmost = True
        dlg.ShowActivated = True
        dlg.WindowStartupLocation = WindowStartupLocation.CenterOwner
        Await Task.Run(Sub() Me.Dispatcher.BeginInvoke(Sub() dlg.ShowDialog()))

        ' 表示処理が完了するまで待機
        Await task1

        dlg.Close()


    End Function

#End Region

#Region "Thumb コントロールのドラッグアンドドロップ移動イベント"

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

    Private Sub Caret_PositionChanged(sender As Object, e As EventArgs)

        Dim texteditor1 = TryCast(sender, TextEditor)
        Dim thumb1 = TryCast(TryCast(TryCast(texteditor1.Parent, DockPanel).Parent, Border).TemplatedParent, ResizableThumb)
        Dim sourceFile = texteditor1.Document.FileName
        Dim offset = texteditor1.TextArea.Caret.Offset

        Dim sourceTree = Me.TreeItems.FirstOrDefault(Function(x) x.FilePath = sourceFile)
        Dim model = Me.CurrentCompilation.GetSemanticModel(sourceTree)
        Dim si = SymbolFinder.FindSymbolAtPositionAsync(model, offset, MSBuildWorkspace.Create()).Result

        If si Is Nothing Then
            Console.WriteLine("not found")
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
        Dim candidateTrees = Me.TreeItems.Where(Function(x) x.GetRoot().DescendantNodes().Any(Function(y) y.ToString() = signature))
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

        ' TODO, コネクタ接続


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

        ' タイトルをセット 
        Dim textblock1 = TryCast(newThumb.Template.FindName("textblock1", newThumb), TextBlock)
        'textblock1.Text = $"{Path.GetFileName(sourceFile)}"
        textblock1.Text = $"{sourceFile}"

        ' ソースをセット
        Dim texteditor1 = TryCast(newThumb.Template.FindName("texteditor1", newThumb), TextEditor)
        texteditor1.Document.Text = File.ReadAllText(sourceFile, EncodeResolver.GetEncoding(sourceFile))
        texteditor1.Document.FileName = sourceFile

        ' メンバー定義位置が見えるようにスクロール
        Dim jumpLine = texteditor1.Document.GetLineByOffset(startLength).LineNumber
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

                                                                   ' 通常のままだと texteditor1 コントロールが取得できないので、イベントハンドラをトラップして、渡してしまう
                                                                   sender = texteditor1
                                                                   Me.Caret_PositionChanged(sender, e)

                                                               End Sub

        ' 移動イベントの購読
        AddHandler newThumb.DragDelta, AddressOf Me.Thumb_DragDelta

        ' キャンバスに登録
        Me.MethodCanvas.Children.Add(newThumb)

        ' TODO, コネクタの接続






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
