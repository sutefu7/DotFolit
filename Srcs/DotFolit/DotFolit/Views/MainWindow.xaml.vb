Imports System.Collections.ObjectModel
Imports System.IO
Imports System.Text
Imports System.Reflection
Imports Microsoft.Win32
Imports System.ComponentModel
Imports System.Windows.Controls.Primitives
Imports Xceed.Wpf.AvalonDock
Imports Xceed.Wpf.AvalonDock.Layout


Class MainWindow

#Region "ウィンドウのドラッグアンドドロップ（sln ファイルの取得）"

    Private Sub Window_PreviewDragOver(sender As Object, e As DragEventArgs)

        If e.Data.GetDataPresent(DataFormats.FileDrop, True) Then
            e.Effects = DragDropEffects.Copy
        Else
            e.Effects = DragDropEffects.None
        End If

        e.Handled = True

    End Sub

    Private Async Sub Window_Drop(sender As Object, e As DragEventArgs)

        Dim slnFiles = TryCast(e.Data.GetData(DataFormats.FileDrop), String())
        If slnFiles Is Nothing Then
            Return
        End If

        Dim slnFile = slnFiles.FirstOrDefault()
        If Path.GetExtension(slnFile).ToLower() <> ".sln" Then
            MessageBox.Show($"{slnFile} は、ソリューションファイルではありません。{vbNewLine}ソリューションファイルを選択してください。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error)
            Return
        End If

        If Not File.Exists(slnFile) Then
            MessageBox.Show($"{slnFile} は、存在しないファイルです。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error)
            Return
        End If

        Await Me.Parse(slnFile)

    End Sub

    Private Async Function Parse(solutionFile As String) As Task

        ' 指定の sln ファイルはビルドされているかチェック
        ' （各プロジェクトにビルド済みアセンブリファイルが存在しているかチェック）
        Dim slnParser = New SolutionParser
        Dim prjParser = New ProjectParser
        Dim projectFiles = slnParser.GetProjectFiles(solutionFile)

        For Each projectFile In projectFiles

            Dim assemblyFile = prjParser.GetAssemblyFile(projectFile)
            If Not File.Exists(assemblyFile) Then

                Me.Activate()
                MessageBox.Show($"{projectFile} プロジェクトのアセンブリファイルが見つかりませんでした。いったんビルドしてから再実行してください。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error)
                Return

            End If

            ' アセンブリファイルを読み込んでおく（アプリケーションドメインに置いておく、後で使う）
            Dim asm = Assembly.LoadFrom(assemblyFile)

        Next

        ' 各ファイルの解析とツリー表示処理は、別スレッドでおこなう
        Dim task1 = Task.Run(Sub()

                                 ' 全プロジェクト配下のソースリストを取得、メモリDBに登録
                                 Dim rosParser = New RoslynParser
                                 Dim nsTable = MemoryDB.Instance.DB.Tables("NamespaceResolution")

                                 For Each projectFile In projectFiles

                                     Dim rootNamespace = prjParser.GetRootNamespace(projectFile)
                                     Dim sourceFiles = prjParser.GetSourceFiles(projectFile)

                                     Dim row = nsTable.NewRow()
                                     row("DefineKind") = "Namespace"
                                     row("ContainerFullName") = String.Empty
                                     row("DefineFullName") = rootNamespace
                                     row("DefineType") = DBNull.Value
                                     row("ReturnType") = DBNull.Value
                                     row("MethodArguments") = DBNull.Value
                                     row("IsPartial") = DBNull.Value
                                     row("IsShared") = DBNull.Value
                                     row("SourceFile") = String.Empty
                                     row("StartLength") = -1
                                     row("EndLength") = -1
                                     row("StartLineNumber") = -1
                                     row("EndLineNumber") = -1
                                     nsTable.Rows.Add(row)

                                     For Each sourceFile In sourceFiles
                                         rosParser.Parse(nsTable, rootNamespace, sourceFile)
                                     Next
                                 Next

                                 ' ソリューションツリーに、モデルをバインドする
                                 Dim treeItems = New ObservableCollection(Of TreeViewItemModel)
                                 Dim treeModel = Me.CreateSolutionTreeData(solutionFile)
                                 treeItems.Add(treeModel)

                                 ' UI スレッド経由で更新
                                 Me.Dispatcher.BeginInvoke(Sub() Me.SolutionTree.ItemsSource = treeItems)

                             End Sub)

        ' 処理中は、進捗状況画面を表示する（進捗状況は、具体的ではなく、応答なしではないことをユーザーに伝える程度）
        Me.Activate()

        Dim dlg = New ProgressWindow
        dlg.Owner = Me
        dlg.Topmost = True
        dlg.ShowActivated = True
        dlg.WindowStartupLocation = WindowStartupLocation.CenterOwner
        Await Task.Run(Sub() Me.Dispatcher.BeginInvoke(Sub() dlg.ShowDialog()))

        ' 解析＆ツリー表示処理が完了するまで待機
        Await task1

        dlg.Close()

    End Function

    Private Function CreateSolutionTreeData(solutionFile As String) As TreeViewItemModel

        ' ソリューションノード
        Dim solutionName = Path.GetFileNameWithoutExtension(solutionFile)
        Dim solutionModel = New TreeViewItemModel With {.Text = solutionName, .FileName = solutionFile, .TreeNodeKind = TreeNodeKinds.SolutionNode, .IsExpanded = True}

        Dim slnParser = New SolutionParser
        Dim prjParser = New ProjectParser
        Dim projectInfos = slnParser.GetProjectDisplayNameAndFiles(solutionFile)
        For Each projectInfo In projectInfos

            ' プロジェクトノード
            Dim displayName = projectInfo.Item1
            Dim projectFile = projectInfo.Item2

            Dim projectModel = New TreeViewItemModel With {.Text = displayName, .FileName = projectFile, .TreeNodeKind = TreeNodeKinds.ProjectNode, .IsExpanded = True}
            solutionModel.AddChild(projectModel)

            ' 参照 dll ノード
            Dim referenceModel = New TreeViewItemModel With {.Text = "参照", .TreeNodeKind = TreeNodeKinds.DependencyNode}
            Dim referenceNames = prjParser.GetReferenceAssemblyNames(projectFile)
            For Each referenceName In referenceNames

                Dim oneModel = New TreeViewItemModel With {.Text = referenceName, .TreeNodeKind = TreeNodeKinds.DependencyNode}
                referenceModel.AddChild(oneModel)

            Next

            projectModel.AddChild(referenceModel)



            ' 自動インポートノード
            Dim importModel = New TreeViewItemModel With {.Text = "自動インポート", .TreeNodeKind = TreeNodeKinds.NamespaceNode}
            Dim importNames = prjParser.GetImportNamespaceNames(projectFile)
            For Each importName In importNames

                Dim oneModel = New TreeViewItemModel With {.Text = importName, .TreeNodeKind = TreeNodeKinds.NamespaceNode}
                importModel.AddChild(oneModel)

            Next

            projectModel.AddChild(importModel)



            ' ソースファイル
            ' Form や Control など、デザイナーファイルとソースファイルのペアの場合がある。この場合、デザイナーファイルをソースファイルの下に登録する
            Dim sourceInfos = prjParser.GetSourceFilesWithDependentUpon(projectFile)
            Dim mainInfos = sourceInfos.Where(Function(x) x.Item2 = String.Empty)
            Dim subInfos = sourceInfos.Where(Function(x) x.Item2 <> String.Empty)

            For Each mainInfo In mainInfos

                Dim sourceFile = mainInfo.Item1
                Dim sourceName = Path.GetFileName(sourceFile)
                Dim mainModel = New TreeViewItemModel With {.Text = sourceName, .FileName = sourceFile, .TreeNodeKind = TreeNodeKinds.SourceNode}

                If subInfos.Any(Function(x) x.Item2 = sourceFile) Then

                    For Each subInfo In subInfos.Where(Function(x) x.Item2 = sourceFile)

                        Dim srcFile = subInfo.Item1
                        Dim srcName = Path.GetFileName(srcFile)
                        Dim subModel = New TreeViewItemModel With {.Text = srcName, .FileName = srcFile, .TreeNodeKind = TreeNodeKinds.GeneratedFileNode}
                        mainModel.AddChild(subModel)

                    Next

                End If

                ' サブフォルダを作成している場合、サブフォルダ数分ノード階層を挟む
                Dim prjDir = Path.GetDirectoryName(projectFile)
                Dim srcDir = Path.GetDirectoryName(sourceFile)
                srcDir = srcDir.Replace(prjDir, String.Empty)

                ' 差分が無ければ、プロジェクトノード直下に登録
                If srcDir = String.Empty Then
                    projectModel.AddChild(mainModel)
                    Continue For
                End If

                Dim subDirs = srcDir.Split(New String() {"\"}, StringSplitOptions.RemoveEmptyEntries)
                Dim parentModel As TreeViewItemModel = Nothing
                Dim currentModel As TreeViewItemModel = Nothing
                Dim i = 1

                ' これから作成しようとしているフォルダが、すでにプロジェクトノードに登録されている場合、そのノードインスタンスを取得、無ければ作成して登録
                If projectModel.Children.Any(Function(x) x.Text = subDirs(0)) Then
                    parentModel = projectModel.Children.FirstOrDefault(Function(x) x.Text = subDirs(0))
                Else
                    parentModel = New TreeViewItemModel With {.Text = subDirs(0), .TreeNodeKind = TreeNodeKinds.FolderNode}
                    projectModel.AddChild(parentModel)
                End If

                ' サブフォルダがある分だけ、繰り返す
                While i < subDirs.Count()

                    If parentModel.Children.Any(Function(x) x.Text = subDirs(i)) Then
                        currentModel = parentModel.Children.FirstOrDefault(Function(x) x.Text = subDirs(i))
                    Else
                        currentModel = New TreeViewItemModel With {.Text = subDirs(i), .TreeNodeKind = TreeNodeKinds.FolderNode}
                        parentModel.AddChild(currentModel)
                    End If

                    ' 現在のフォルダを親フォルダに変えて、再帰
                    parentModel = currentModel
                    i += 1

                End While

                parentModel.AddChild(mainModel)

            Next

        Next

        Return solutionModel

    End Function

    'Private Iterator Function CreateSolutionTreeDataInternal(solutionFile As String) As IEnumerable(Of TreeViewItemModel)

    'End Function

#End Region

#Region "メニューのファイルを開く（sln ファイルの取得）"

    Private Async Sub MenuItem_Click(sender As Object, e As RoutedEventArgs)

        Dim dlg = New OpenFileDialog
        dlg.Title = "ファイルの選択"
        dlg.FileName = "*.sln"
        dlg.Filter = "ソリューション ファイル(*.sln)|*.sln"

        Dim response = dlg.ShowDialog()
        If response.HasValue AndAlso response.Value Then

            Dim slnFile = dlg.FileName
            Await Me.Parse(slnFile)

        End If

    End Sub

#End Region

#Region "ソリューションエクスプローラーのノードクリック"

    Private Sub SolutionTree_SelectedItemChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Object))

        ' ソースタブを選択した際に、対応するノードを強制選択した場合、抑制フラグがオンになる
        ' この場合、何もしない
        If Me.dontWork Then
            Return
        End If

        Dim selectedModel = TryCast(e.NewValue, TreeViewItemModel)
        If selectedModel Is Nothing Then
            Return
        End If

        If selectedModel.TreeNodeKind <> TreeNodeKinds.SourceNode AndAlso selectedModel.TreeNodeKind <> TreeNodeKinds.GeneratedFileNode Then
            Return
        End If

        Dim sourceFile = selectedModel.FileName
        Me.AddSourcePane(sourceFile)

    End Sub

    Private Sub AddSourcePane(sourceFile As String)

        ' すでに表示中の場合、エディタタブをアクティブする
        Dim contentId = sourceFile
        If (Me.SourcePane.ChildrenCount <> 0) AndAlso (Me.SourcePane.Children.Any(Function(x) x.ContentId = contentId)) Then

            ' ソリューションツリー選択時、既存表示されたソースタブを選択すると、ソリューションツリー選択イベントを発生させてしまうため、抑制させてから選択する
            Dim tabPage = Me.SourcePane.Children.FirstOrDefault(Function(x) x.ContentId = contentId)
            RemoveHandler tabPage.IsSelectedChanged, AddressOf Me.LayoutDocument_IsSelectedChanged
            tabPage.IsSelected = True
            AddHandler tabPage.IsSelectedChanged, AddressOf Me.LayoutDocument_IsSelectedChanged

            Return

        End If

        ' 見つからなかった場合、新規登録
        Dim sourceName = Path.GetFileName(sourceFile)
        Dim newPage = New LayoutDocument With {.Title = sourceName, .ContentId = sourceFile}
        newPage.IsSelected = True
        newPage.IsActive = True
        AddHandler newPage.IsSelectedChanged, AddressOf Me.LayoutDocument_IsSelectedChanged

        Dim newContent = New EditorUserControl
        newContent.InitializeData(sourceFile)
        newPage.Content = newContent

        Me.SourcePane.Children.Add(newPage)

    End Sub

#End Region

#Region "ソースタブのアクティブ変更"

    ' ソースタブを選択したら、対応するソリューションツリーのノードを選択させたい
    ' ただし、強制選択したら、ノード選択イベントが発生してしまうため、フラグで制御する
    Private dontWork As Boolean = False

    ' 選択された、または未選択になった場合に発生
    Private Sub LayoutDocument_IsSelectedChanged(sender As Object, e As EventArgs)

        Dim tabPage = TryCast(sender, LayoutDocument)
        If (tabPage Is Nothing) OrElse (tabPage.IsSelected = False) Then
            Return
        End If

        ' フラグをオン
        Me.dontWork = True

        ' ソースタブを一意に識別できる ContentId = SourceFile なので、ここから取得
        Dim sourceFile = tabPage.ContentId
        Dim isFound = False

        For Each model As TreeViewItemModel In Me.SolutionTree.Items
            Me.SelectedItemChangedFromTabPage(model, sourceFile, isFound)
        Next

        ' TODO
        ' アクティブなソースタブを切り替えた際、表示中の継承関係図、フローチャート図が切り替わらない
        ' （先ほどまで表示していたクラス用の図形のまま残る、切り替え後のクラス用の図形に切り替わらない）
        'Dim editor = TryCast(tabPage.Content, EditorUserControl)
        'editor.Caret_PositionChanged(Me, EventArgs.Empty) ' Private だから呼び出せない
        'editor.texteditor1.TextArea.Caret.Offset = editor.texteditor1.TextArea.Caret.Offset

        ' フラグ解除
        Me.dontWork = False

    End Sub

    Private Sub SelectedItemChangedFromTabPage(model As TreeViewItemModel, sourceFile As String, ByRef isFound As Boolean)

        ' 他のところで見つけたのであれば、何もしないで抜ける
        If isFound Then
            Return
        End If

        If model.FileName = sourceFile Then

            model.IsSelected = True
            isFound = True
            Return

        End If

        If model.Children.Any() Then
            For Each child In model.Children
                Me.SelectedItemChangedFromTabPage(child, sourceFile, isFound)
            Next
        End If

    End Sub

#End Region

#Region "継承関係図キャンバスのマウスホイール変更（ズームイン、アウト）"

    ' [WPF]コントロールキーあるいはシフトキーが押されているかどうかを取得する
    ' https://gist.github.com/pinzolo/3080481


    Private Sub InheritsCanvas_MouseWheel(sender As Object, e As MouseWheelEventArgs)

        Dim isDownLeftControlKey = ((Keyboard.GetKeyStates(Key.LeftCtrl) And KeyStates.Down) = KeyStates.Down)
        Dim isDownRightControlKey = ((Keyboard.GetKeyStates(Key.RightCtrl) And KeyStates.Down) = KeyStates.Down)
        Dim isDownControlKey = isDownLeftControlKey OrElse isDownRightControlKey

        If isDownControlKey Then

            If 0 < e.Delta Then

                InheritsScaleTransform.ScaleX *= 1.1
                InheritsScaleTransform.ScaleY *= 1.1

            Else

                InheritsScaleTransform.ScaleX /= 1.1
                InheritsScaleTransform.ScaleY /= 1.1

            End If

        End If

    End Sub

#End Region

#Region "メソッドフローチャートキャンバスのマウスホイール変更（ズームイン、アウト）"

    ' [WPF]コントロールキーあるいはシフトキーが押されているかどうかを取得する
    ' https://gist.github.com/pinzolo/3080481


    Private Sub FlowChartCanvas_MouseWheel(sender As Object, e As MouseWheelEventArgs)

        Dim isDownLeftControlKey = ((Keyboard.GetKeyStates(Key.LeftCtrl) And KeyStates.Down) = KeyStates.Down)
        Dim isDownRightControlKey = ((Keyboard.GetKeyStates(Key.RightCtrl) And KeyStates.Down) = KeyStates.Down)
        Dim isDownControlKey = isDownLeftControlKey OrElse isDownRightControlKey

        If isDownControlKey Then

            If 0 < e.Delta Then

                FlowChartScaleTransform.ScaleX *= 1.1
                FlowChartScaleTransform.ScaleY *= 1.1

            Else

                FlowChartScaleTransform.ScaleX /= 1.1
                FlowChartScaleTransform.ScaleY /= 1.1

            End If

        End If

    End Sub

#End Region

#Region "メソッドの追跡...コンテキストメニューのクリック"

    Private Sub TreeViewItemMenuItem_Click(sender As Object, e As RoutedEventArgs)

        Dim model = TryCast(Me.SolutionTree.SelectedItem, TreeViewItemModel)
        If model Is Nothing Then
            Return
        End If

        Dim solutionModel As TreeViewItemModel = model.Parent
        While True

            If solutionModel.TreeNodeKind = TreeNodeKinds.SolutionNode Then
                Exit While
            End If
            solutionModel = solutionModel.Parent

        End While

        Dim dlg = New MethodWindow
        dlg.SolutionModel = solutionModel
        dlg.SourceModel = model

        Dim parentWindow = TryCast(Window.GetWindow(Me), MainWindow)
        If parentWindow IsNot Nothing Then
            dlg.Owner = parentWindow
            dlg.WindowStartupLocation = WindowStartupLocation.CenterOwner
        End If

        dlg.ShowDialog()
        dlg = Nothing

    End Sub

#End Region

#Region ""

#End Region

#Region ""

#End Region

#Region ""

#End Region

End Class
