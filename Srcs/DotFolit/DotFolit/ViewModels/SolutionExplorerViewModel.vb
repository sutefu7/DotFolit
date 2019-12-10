Imports System.Collections.ObjectModel
Imports System.IO
Imports System.Reflection
Imports Livet
Imports Livet.Messaging
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.MSBuild
Imports Microsoft.CodeAnalysis.VisualBasic


Public Class SolutionExplorerViewModel
    Inherits AnchorablePaneViewModel

#Region "フィールド、プロパティ"

    Private SolutionFile As String = String.Empty

#End Region

#Region "AnchorablePane 用のプロパティ"

    Public Overrides ReadOnly Property Title As String
        Get
            Return "ソリューション エクスプローラー"
        End Get
    End Property

    Public Overrides ReadOnly Property ContentId As String
        Get
            Return "SolutionExplorer"
        End Get
    End Property

#End Region

#Region "TreeItems変更通知プロパティ"
    Private _TreeItems As ObservableCollection(Of TreeViewItemModel)

    Public Property TreeItems() As ObservableCollection(Of TreeViewItemModel)
        Get
            Return _TreeItems
        End Get
        Set(ByVal value As ObservableCollection(Of TreeViewItemModel))
            If (value Is Nothing) Then Return
            _TreeItems = value
            RaisePropertyChanged()
        End Set
    End Property
#End Region

#Region "SelectedNode変更通知プロパティ"
    Private _SelectedNode As TreeViewItemModel

    Public Property SelectedNode() As TreeViewItemModel
        Get
            Return _SelectedNode
        End Get
        Set(ByVal value As TreeViewItemModel)
            If (value Is Nothing) Then Return
            _SelectedNode = value
            RaisePropertyChanged()
        End Set
    End Property
#End Region

#Region "コンストラクタ"

    Public Sub New()

        ' 非同期スレッドからバインドデータにアクセスできるようにする
        Me._TreeItems = New ObservableCollection(Of TreeViewItemModel)
        BindingOperations.EnableCollectionSynchronization(Me.TreeItems, New Object())

    End Sub

#End Region

#Region "ソリューションエクスプローラーツリーのデータ表示"

    ' TreeView バインド用の VM を用意
    Public Async Sub ShowData(solutionFile As String)

        Using vm = New ProgressViewModel

            Me.SolutionFile = solutionFile
            vm.DoWork = AddressOf Me.DoWork
            vm.IsDoWorkFinishedAfterWait = True
            Await Messenger.RaiseAsync(New TransitionMessage(vm, "ShowProgressView"))

        End Using

    End Sub

    ' 進捗画面側で実行してもらう処理
    Private Async Sub DoWork()

        Dim ws = MSBuildWorkspace.Create()
        Dim solutionInfo = Await ws.OpenSolutionAsync(SolutionFile)

        ' 事前チェック、各プロジェクトにアセンブリファイルが生成されていること（リビルドを１回以上していること）
        For Each projectInfo In solutionInfo.Projects

            ' アセンブリファイル
            Dim outputFilePath = GetOutputFilePath(projectInfo)
            If Not File.Exists(outputFilePath) Then

                Dim message = $"{projectInfo.Name} プロジェクトのアセンブリファイルが見つかりませんでした。{Environment.NewLine}"
                message &= $"{outputFilePath}{Environment.NewLine}"
                message &= $"{projectInfo.Name} プロジェクトがビルドされていない可能性があります。お手数ですが、いったんビルドしてから、再度ご利用ください。"

                Await Me.ShowErrorMessageAsync(message)
                Return

            End If

        Next

        ' ソリューションエクスプローラーのツリー表示用に、バインドデータの準備
        Dim task1 = Task.Run(Sub()

                                 Dim treeModel = Me.CreateTreeData(solutionInfo)
                                 Me.TreeItems = New ObservableCollection(Of TreeViewItemModel) From {treeModel}

                             End Sub)

        ' メソッド追跡画面用に、変数準備
        Dim task2 = Task.Run(Sub()

                                 Dim parser = New ProjectParser
                                 Dim dllItems = New List(Of MetadataReference)
                                 Dim srcItems = New List(Of SyntaxTree)
                                 Dim srcItems2 = New List(Of SyntaxTree)
                                 Dim compItems = New List(Of VisualBasicCompilation)

                                 Dim mscorlib = MetadataReference.CreateFromFile(GetType(Object).GetTypeInfo().Assembly.Location)
                                 dllItems.Add(mscorlib)

                                 For Each projectInfo In solutionInfo.Projects

                                     ' 参照dll
                                     'For Each metaInfo In projectInfo.MetadataReferences

                                     '    If Not dllItems.Any(Function(x) x.Display = metaInfo.Display) Then
                                     '        dllItems.Add(MetadataReference.CreateFromFile(metaInfo.Display))
                                     '    End If

                                     'Next

                                     ' TODO, 名前だけなので、GAC パスや Nuget, その他個別パスを探して、フルパス変換する必要あり
                                     'Dim metadataReferences = parser.GetReferenceAssemblyNames(projectInfo.FilePath)
                                     'For Each metaInfo In metadataReferences

                                     '    Dim location = System.Type.GetType(metaInfo).Assembly.Location ' NG. 名前空間だけではなく、何かのクラスもくっつけないとダメ

                                     '    If Not dllItems.Any(Function(x) x.Display = location) Then
                                     '        dllItems.Add(MetadataReference.CreateFromFile(location))
                                     '    End If

                                     'Next

                                     ' ソースファイル
                                     'For Each sourceInfo In projectInfo.Documents

                                     '    Dim source = File.ReadAllText(sourceInfo.FilePath, EncodeResolver.GetEncoding(sourceInfo.FilePath))
                                     '    Dim tree = VisualBasicSyntaxTree.ParseText(source,, sourceInfo.FilePath)
                                     '    srcItems.Add(tree)

                                     'Next

                                     srcItems2.Clear()
                                     Dim documents = parser.GetSourceFiles(projectInfo.FilePath)
                                     For Each sourceInfo In documents

                                         Dim source = File.ReadAllText(sourceInfo, EncodeResolver.GetEncoding(sourceInfo))
                                         Dim tree = VisualBasicSyntaxTree.ParseText(source,, sourceInfo)
                                         srcItems2.Add(tree)

                                     Next



                                     ' アセンブリファイル
                                     Dim outputFilePath = GetOutputFilePath(projectInfo)
                                     If Not dllItems.Any(Function(x) x.Display = outputFilePath) Then
                                         'dllItems.Add(MetadataReference.CreateFromFile(outputFilePath))
                                     End If

                                     ' クラスの継承関係図の準備で必要なため、読み込んでおく（アプリケーションドメインに置いておく）
                                     Dim asm = Assembly.LoadFrom(outputFilePath)

                                     Dim options = New VisualBasicCompilationOptions(
                                        outputKind:=If(Path.GetExtension(outputFilePath).ToLower() = ".exe", OutputKind.ConsoleApplication, OutputKind.DynamicallyLinkedLibrary),
                                        rootNamespace:=parser.GetRootNamespace(projectInfo.FilePath))

                                     ' メインプロジェクトのソースを追加した後で、登録済みのソースを追加する（検索優先度の調整？）
                                     srcItems2.AddRange(srcItems)
                                     Dim compilation = VisualBasicCompilation.Create(
                                        projectInfo.AssemblyName,
                                        srcItems2,
                                        dllItems,
                                        options)

                                     ' メイン処理用のソースリストは、登録順に直してから追加する
                                     srcItems.Clear()
                                     srcItems.AddRange(srcItems2)
                                     srcItems.Reverse()
                                     compItems.Add(compilation)

                                 Next


                                 ' メモリDBにセット
                                 MemoryDB.Instance.SyntaxTreeItems = srcItems
                                 MemoryDB.Instance.CompilationItems = compItems

                             End Sub)

        ' メモリDB用に、ソース内容解析と登録
        Dim task3 = Task.Run(Sub()

                                 Dim parser = New ProjectParser
                                 For Each projectInfo In solutionInfo.Projects

                                     ' ソースファイル
                                     Dim projectNamespace = parser.GetRootNamespace(projectInfo.FilePath)
                                     'For Each sourceInfo In projectInfo.Documents

                                     '    Dim source = File.ReadAllText(sourceInfo.FilePath, EncodeResolver.GetEncoding(sourceInfo.FilePath))
                                     '    Dim tree = VisualBasicSyntaxTree.ParseText(source,, sourceInfo.FilePath)
                                     '    Dim walker = New SourceSyntaxWalker
                                     '    walker.Parse(MemoryDB.Instance.DB.Tables("NamespaceResolution"), projectNamespace, sourceInfo.FilePath, tree)

                                     'Next


                                     Dim documents = parser.GetSourceFiles(projectInfo.FilePath)
                                     For Each sourceInfo In documents

                                         Dim source = File.ReadAllText(sourceInfo, EncodeResolver.GetEncoding(sourceInfo))
                                         Dim tree = VisualBasicSyntaxTree.ParseText(source,, sourceInfo)
                                         Dim walker = New SourceSyntaxWalker
                                         walker.Parse(MemoryDB.Instance.DB.Tables("NamespaceResolution"), projectNamespace, sourceInfo, tree)

                                     Next

                                 Next

                             End Sub)

        Await Task.WhenAll(task1, task2, task3)

    End Sub

    Private Function GetOutputFilePath(projectInfo As Project) As String

        If Not String.IsNullOrWhiteSpace(projectInfo.OutputFilePath) Then
            Return projectInfo.OutputFilePath
        End If

        Dim projectDir = Path.GetDirectoryName(projectInfo.FilePath)
        Dim exeDir = Path.Combine(projectDir, "bin", "Debug")

        If Not Directory.Exists(exeDir) Then
            exeDir = Path.Combine(projectDir, "bin", "Release")
        End If

        If Not Directory.Exists(exeDir) Then
            Throw New DirectoryNotFoundException("ビルド済みフォルダが見つかりませんでした。一度ビルドを実行してください。")
        End If

        Dim outputFilePath = Directory.EnumerateFiles(exeDir, $"{projectInfo.AssemblyName}.*").
            Where(Function(x) Path.GetExtension(x).ToLower() = ".exe" OrElse Path.GetExtension(x).ToLower() = ".dll").
            FirstOrDefault()

        If String.IsNullOrWhiteSpace(outputFilePath) OrElse Not File.Exists(outputFilePath) Then
            Throw New FileNotFoundException("アセンブリファイルが見つかりませんでした。一度ビルドを実行してください。")
        End If

        Return outputFilePath

    End Function

    Private Function CreateTreeData(solutionInfo As Solution) As TreeViewItemModel

        ' ソリューションノード
        Dim solutionName = Path.GetFileNameWithoutExtension(solutionInfo.FilePath)
        Dim solutionModel = New TreeViewItemModel With {.Text = solutionName, .FileName = solutionInfo.FilePath, .TreeNodeKind = TreeNodeKinds.SolutionNode, .IsExpanded = True}
        Dim parser = New ProjectParser

        For Each projectInfo In solutionInfo.Projects

            ' プロジェクトノード
            Dim projectModel = New TreeViewItemModel With {.Text = projectInfo.Name, .FileName = projectInfo.FilePath, .TreeNodeKind = TreeNodeKinds.ProjectNode, .IsExpanded = True}
            solutionModel.Children.Add(projectModel)

            ' 参照 dll ノード
            Dim referenceModel = New TreeViewItemModel With {.Text = "参照", .TreeNodeKind = TreeNodeKinds.DependencyNode}
            Dim referenceNames = parser.GetReferenceAssemblyNames(projectInfo.FilePath)
            For Each referenceName In referenceNames

                Dim oneModel = New TreeViewItemModel With {.Text = referenceName, .TreeNodeKind = TreeNodeKinds.DependencyNode}
                referenceModel.Children.Add(oneModel)

            Next
            projectModel.Children.Add(referenceModel)

            ' 自動インポートノード
            Dim importModel = New TreeViewItemModel With {.Text = "自動インポート", .TreeNodeKind = TreeNodeKinds.NamespaceNode}
            Dim importNames = parser.GetImportNamespaceNames(projectInfo.FilePath)
            For Each importName In importNames

                Dim oneModel = New TreeViewItemModel With {.Text = importName, .TreeNodeKind = TreeNodeKinds.NamespaceNode}
                importModel.Children.Add(oneModel)

            Next
            projectModel.Children.Add(importModel)

            ' ソースファイル
            ' Form や Control など、デザイナーファイルとソースファイルのペアの場合がある。この場合、デザイナーファイルをソースファイルの下に登録する
            Dim sourceInfos = parser.GetSourceFilesWithDependentUpon(projectInfo.FilePath)
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
                        mainModel.Children.Add(subModel)

                    Next

                End If

                ' サブフォルダを作成している場合、サブフォルダ数分ノード階層を挟む
                Dim prjDir = Path.GetDirectoryName(projectInfo.FilePath)
                Dim srcDir = Path.GetDirectoryName(sourceFile)
                srcDir = srcDir.Replace(prjDir, String.Empty)

                ' 差分が無ければ、プロジェクトノード直下に登録
                If srcDir = String.Empty Then
                    projectModel.Children.Add(mainModel)
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
                    projectModel.Children.Add(parentModel)
                End If

                ' サブフォルダがある分だけ、繰り返す
                While i < subDirs.Count()

                    If parentModel.Children.Any(Function(x) x.Text = subDirs(i)) Then
                        currentModel = parentModel.Children.FirstOrDefault(Function(x) x.Text = subDirs(i))
                    Else
                        currentModel = New TreeViewItemModel With {.Text = subDirs(i), .TreeNodeKind = TreeNodeKinds.FolderNode}
                        parentModel.Children.Add(currentModel)
                    End If

                    ' 現在のフォルダを親フォルダに変えて、再帰
                    parentModel = currentModel
                    i += 1

                End While

                parentModel.Children.Add(mainModel)

            Next

        Next

        Return solutionModel

    End Function

#End Region

#Region "ソリューションエクスプローラーツリーのノード選択"

    Public Sub SolutionTree_SelectedItemChanged(e As TreeViewItemModel)

        ' 値を更新して、プロパティ変更通知を発生させる（MainViewModel 側で受け取る）
        Me.SelectedNode = e

    End Sub

#End Region

#Region "ソリューションエクスプローラーツリーの右クリック→コンテキストメニュー処理（メソッドの追跡）"

    Public Async Sub MethodSearchMenuItem_Click(e As TreeViewItemModel)

        Using vm = New MethodViewModel

            vm.SelectedNode = e
            Await Messenger.RaiseAsync(New TransitionMessage(vm, "ShowMethodView"))

        End Using

    End Sub

#End Region

End Class
