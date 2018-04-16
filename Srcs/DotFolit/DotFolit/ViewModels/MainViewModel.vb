Imports System.ComponentModel
Imports System.Collections.ObjectModel
Imports System.IO
Imports System.Text
Imports Livet
Imports Livet.Messaging.IO
Imports Livet.Messaging.Windows
Imports Livet.EventListeners
Imports Livet.EventListeners.WeakEvents
Imports GongSolutions.Wpf.DragDrop
Imports Livet.Commands
Imports Xceed.Wpf.AvalonDock


Public Class MainViewModel
    Inherits ViewModel
    Implements IDropTarget

#Region "SolutionExplorerVM変更通知プロパティ"
    Private _SolutionExplorerVM As SolutionExplorerViewModel = Nothing

    Public Property SolutionExplorerVM() As SolutionExplorerViewModel
        Get
            Return _SolutionExplorerVM
        End Get
        Set(ByVal value As SolutionExplorerViewModel)
            If (value Is Nothing) Then Return
            _SolutionExplorerVM = value
            RaisePropertyChanged()
        End Set
    End Property
#End Region

#Region "Anchorables変更通知プロパティ"
    Private _Anchorables As ObservableCollection(Of ViewModel) = Nothing

    Public Property Anchorables() As ObservableCollection(Of ViewModel)
        Get
            Return _Anchorables
        End Get
        Set(ByVal value As ObservableCollection(Of ViewModel))
            If (value Is Nothing) Then Return
            _Anchorables = value
            RaisePropertyChanged()
        End Set
    End Property
#End Region

#Region "Documents変更通知プロパティ"
    Private _Documents As ObservableCollection(Of ViewModel) = Nothing

    Public Property Documents() As ObservableCollection(Of ViewModel)
        Get
            Return _Documents
        End Get
        Set(ByVal value As ObservableCollection(Of ViewModel))
            If (value Is Nothing) Then Return
            _Documents = value
            RaisePropertyChanged()
        End Set
    End Property
#End Region

#Region "コンストラクタ"

    Public Sub New()

        ' 画面デザイン時は処理しない（これにより画面デザインがプレビュー反映されなくなったが、保留）
        If DesignerProperties.GetIsInDesignMode(New DependencyObject()) Then
            Return
        End If

        Me.SolutionExplorerVM = New SolutionExplorerViewModel
        Me.Anchorables = New ObservableCollection(Of ViewModel)
        Me.Anchorables.Add(SolutionExplorerVM)

        Me.Documents = New ObservableCollection(Of ViewModel)

        ' ソリューションエクスプローラーペインからの、プロパティ変更通知を受け取る
        ' （ソリューションエクスプローラーツリー内のソースノードを、クリックした際のイベント処理を引き継ぐ）
        Dim listener = New PropertyChangedEventListener(Me.SolutionExplorerVM)
        listener.RegisterHandler(AddressOf Me.SolutionExplorerVM_PropertyChanged)
        Me.CompositeDisposable.Add(listener)

    End Sub

#End Region

#Region "IDropTarget メンバー（ソリューションファイルのドラッグアンドドロップ）"

    Public Sub DragOver(dropInfo As IDropInfo) Implements IDropTarget.DragOver

        Dim dropData = TryCast(dropInfo.Data, DataObject)
        Dim dropFile = dropData.GetFileDropList().Cast(Of String)().FirstOrDefault()

        If Path.GetExtension(dropFile).ToLower() = ".sln" Then
            dropInfo.Effects = DragDropEffects.Copy
        Else
            dropInfo.Effects = DragDropEffects.None
        End If

    End Sub

    Public Async Sub Drop(dropInfo As IDropInfo) Implements IDropTarget.Drop

        Dim dropData = TryCast(dropInfo.Data, DataObject)
        Dim dropFile = dropData.GetFileDropList().Cast(Of String)().FirstOrDefault()

        If Path.GetExtension(dropFile).ToLower() <> ".sln" Then
            Return
        End If

        ' エクスプローラー画面から D&D した場合、自画面が非アクティブ状態のままとなるので、自画面をアクティブに切り替える
        ' こうしないと、後続処理の進捗画面が表示されない現象が発生してしまう（SolutionExplorerVM 側で対策をおこなってもいいのかも）
        Await Messenger.RaiseAsync(New WindowActionMessage("WindowActiveAction"))

        Me.SolutionExplorerVM.ShowData(dropFile)

    End Sub

#End Region

#Region "ファイルを開くダイアログの戻り値（ソリューションファイルの選択メニューから）"

    ' ファイルを開くダイアログの選択結果コールバック
    Public Sub OpenFileDialogCallback(result As OpeningFileSelectionMessage)

        If result.Response Is Nothing Then
            Return
        End If

        Dim solutionFile = result.Response.FirstOrDefault()
        Me.SolutionExplorerVM.ShowData(solutionFile)

    End Sub

#End Region

#Region "ソリューションエクスプローラーペイン内のイベント処理"

    Private Sub SolutionExplorerVM_PropertyChanged(sender As Object, e As PropertyChangedEventArgs)

        ' ソリューションエクスプローラーのツリーノード選択
        If e.PropertyName = NameOf(SolutionExplorerVM.SelectedNode) Then

            Dim selectedModel = SolutionExplorerVM.SelectedNode
            Select Case selectedModel.TreeNodeKind

                Case TreeNodeKinds.SolutionNode
                    ' ソリューションノードをクリックした
                    ' 各プロジェクト別に、参照関係図の表示する？

                Case TreeNodeKinds.ProjectNode
                    ' プロジェクトノードをクリックした
                    ' 該当プロジェクトの参照関係図の表示する？

                Case TreeNodeKinds.SourceNode, TreeNodeKinds.GeneratedFileNode
                    ' ソースノードをクリックした

                    Dim foundModel = Me.Documents.OfType(Of SourceViewModel)().FirstOrDefault(Function(x) x.SourceFile = selectedModel.FileName)
                    If foundModel IsNot Nothing Then
                        ' すでに登録済み（表示済み）なので、再アクティブに切り替える
                        foundModel.IsSelected = True
                    Else
                        ' 新規登録して表示する
                        Me.AddSourcePane(selectedModel)
                    End If

            End Select


        End If

    End Sub

    Private Sub AddSourcePane(e As TreeViewItemModel)

        Dim vm = New SourceViewModel
        vm.SourceFile = e.FileName
        vm.SourceCode = File.ReadAllText(e.FileName, EncodeResolver.GetEncoding(e.FileName))
        vm.IsSelected = True

        Me.Documents.Add(vm)

    End Sub

#End Region

#Region "DocumentClosingCommand"

    ' イベント引数を必要とするため、
    ' イベント発生をビヘイビアで捉えて、バインドしているコマンド経由でイベント引数を受け取る

    ' ※エディタペインがアクティブではなくても×ボタンを押せるため、ActivePain みたいな依存関係プロパティを利用した削除処理では、対応できないと考えた
    ' そういうプロパティは無いけど

    Private _DocumentClosingCommand As ListenerCommand(Of DocumentClosingEventArgs)

    Public ReadOnly Property DocumentClosingCommand() As ListenerCommand(Of DocumentClosingEventArgs)
        Get
            If _DocumentClosingCommand Is Nothing Then
                _DocumentClosingCommand = New ListenerCommand(Of DocumentClosingEventArgs)(AddressOf DocumentClosing)
            End If
            Return _DocumentClosingCommand
        End Get
    End Property

    Private Sub DocumentClosing(ByVal parameter As DocumentClosingEventArgs)

        Dim vm = TryCast(parameter.Document.Content, SourceViewModel)
        If vm Is Nothing Then
            Return
        End If

        Me.Documents.Remove(vm)

        ' 内部的には View のインスタンスを再利用しているのか？
        ' View が破棄されると例外エラー発生してしまう現象の対応
        ' 対応として、キャンセルフラグを立てた状態で戻す

        ' ただし、上記の命令通り、ViewModel 側は削除する
        ' SolutionExplorerVM_PropertyChanged イベント処理内で、表示中の判定として扱われてしまうため

        parameter.Cancel = True

        ' System.NullReferenceException はハンドルされませんでした。
        ' Message: 型 'System.NullReferenceException' のハンドルされていない例外が Xceed.Wpf.AvalonDock.dll で発生しました
        ' 追加情報:オブジェクト参照がオブジェクト インスタンスに設定されていません。

        ' Correctly handling document-close and tool-hide in a WPF app with AvalonDock+Caliburn Micro
        ' https://stackoverflow.com/questions/28194046/correctly-handling-document-close-and-tool-hide-in-a-wpf-app-with-avalondockcal
        ' 
        ' ※こちらは参考程度
        ' WPF - AvalonDock - Closing Document
        ' https://stackoverflow.com/questions/18359818/wpf-avalondock-closing-document
        ' 
        ' MVVM Passing EventArgs As Command Parameter
        ' https://stackoverflow.com/questions/6205472/mvvm-passing-eventargs-as-command-parameter



    End Sub
#End Region

End Class
