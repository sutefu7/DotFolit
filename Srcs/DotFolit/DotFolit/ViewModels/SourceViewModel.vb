Imports System.Collections.ObjectModel
Imports System.Data
Imports System.IO
Imports Livet
Imports ICSharpCode.AvalonEdit
Imports ICSharpCode.AvalonEdit.Document
Imports ICSharpCode.AvalonEdit.Editing
Imports ICSharpCode.AvalonEdit.Folding
Imports Livet.Commands


Public Class SourceViewModel
    Inherits DocumentPaneViewModel

    ' AvalonDock 関連

#Region "DocumentPane 用のプロパティ"

    Public Overrides ReadOnly Property Title As String
        Get
            Return Path.GetFileName(Me.SourceFile)
        End Get
    End Property

    Public Overrides ReadOnly Property ContentId As String
        Get
            Return Me.SourceFile
        End Get
    End Property

#End Region


    ' AvalonEdit 関連

#Region "SourceFile変更通知プロパティ"
    Private _SourceFile As String

    Public Property SourceFile() As String
        Get
            Return _SourceFile
        End Get
        Set(ByVal value As String)
            If (_SourceFile = value) Then Return
            _SourceFile = value
            RaisePropertyChanged()
        End Set
    End Property
#End Region
#Region "SourceCode変更通知プロパティ"
    Private _SourceCode As String

    Public Property SourceCode() As String
        Get
            Return _SourceCode
        End Get
        Set(ByVal value As String)
            If (_SourceCode = value) Then Return
            _SourceCode = value
            RaisePropertyChanged()
        End Set
    End Property
#End Region
#Region "CaretLocation変更通知プロパティ"

    ' キャレットの、行・列位置。５行目１２文字目など

    Private _CaretLocation As TextLocation

    Public Property CaretLocation() As TextLocation
        Get
            Return _CaretLocation
        End Get
        Set(ByVal value As TextLocation)
            If (_CaretLocation = value) Then Return
            _CaretLocation = value
            RaisePropertyChanged()
        End Set
    End Property
#End Region
#Region "CaretOffset変更通知プロパティ"

    ' キャレットの文字数位置。
    ' ソースコードを１つの String として見た際、何文字目か

    Private _CaretOffset As Integer

    Public Property CaretOffset() As Integer
        Get
            Return _CaretOffset
        End Get
        Set(ByVal value As Integer)
            If (_CaretOffset = value) Then Return
            _CaretOffset = value
            RaisePropertyChanged()
        End Set
    End Property
#End Region


    ' クラスのメンバーツリー関連

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


    ' クラス継承関係図 関連

#Region "InheritsItems変更通知プロパティ"
    Private _InheritsItems As ObservableCollection(Of InheritsItemModel)

    Public Property InheritsItems() As ObservableCollection(Of InheritsItemModel)
        Get
            Return _InheritsItems
        End Get
        Set(ByVal value As ObservableCollection(Of InheritsItemModel))
            If (value Is Nothing) Then Return
            _InheritsItems = value
            RaisePropertyChanged()
        End Set
    End Property
#End Region


    ' メソッドのフローチャート図 関連

#Region "MethodModel変更通知プロパティ"
    Private _MethodModel As MethodTemplateModel

    Public Property MethodModel() As MethodTemplateModel
        Get
            Return _MethodModel
        End Get
        Set(ByVal value As MethodTemplateModel)
            If (value Is Nothing) Then Return
            _MethodModel = value
            RaisePropertyChanged()
        End Set
    End Property
#End Region


    ' 内部処理用

#Region "フィールド、プロパティ"

    ' それぞれ、同じ NamespaceResolution テーブル用のデータビューだが、
    ' 複数メソッド内から同時アクセスで扱いたいため、機能ごとに３つ分用意して使う
    Private nsViewForInherits As DataView = Nothing
    Private nsViewForMember As DataView = Nothing
    Private nsViewForMethod As DataView = Nothing

#End Region

#Region "コンストラクタ"

    Public Sub New()

        Me.CanClose = True

        Me._SourceFile = String.Empty
        Me._SourceCode = String.Empty
        Me._CaretLocation = New TextLocation
        Me._CaretOffset = 0

        Me._TreeItems = Nothing
        Me._InheritsItems = Nothing
        Me._MethodModel = Nothing

        Me.InheritsItems = New ObservableCollection(Of InheritsItemModel)

    End Sub

#End Region

#Region "エディタ内にあるキャレットカーソル位置の移動イベント"

    Public Sub BindableTextEditor_CaretPositionChanged()

        ' キャレット位置の範囲に該当するクラスやメソッドが無いか、メモリDBを見ながら探す
        ' クラスが該当したら、継承関係図、メンバーツリーの表示
        ' メソッドが該当したら、フローチャート図の表示

        If Me.nsViewForInherits Is Nothing Then
            Dim nsTable = MemoryDB.Instance.DB.Tables("NamespaceResolution")
            Me.nsViewForInherits = nsTable.AsDataView()
            Me.nsViewForMember = nsTable.AsDataView()
            Me.nsViewForMethod = nsTable.AsDataView()
        End If

        ' クラスの継承関係図
        Me.ShowClassInheritsTree()

        ' クラスのメンバーツリー
        Me.ShowClassMemberTree()

        ' メソッドのフローチャート図
        Me.ShowMethodFlowchart()

    End Sub

    Private Sub ShowClassInheritsTree()

        Dim nsView = Me.nsViewForInherits
        nsView.RowFilter = $"SourceFile='{Me.SourceFile}' AND DefineKind IN ('Class', 'Structure', 'Interface', 'Module', 'Enum', 'Delegate') "
        nsView.Sort = "StartLength, EndLength"

        If nsView.Count = 0 Then
            Me.InheritsItems = New ObservableCollection(Of InheritsItemModel)
            Return
        End If

        Dim defineFullName = String.Empty
        Dim defineKind = String.Empty

        ' クラス定義行の範囲にいるかどうか
        If Not Me.IsPositionInnerMember(nsView, defineFullName, defineKind) Then
            Me.InheritsItems = New ObservableCollection(Of InheritsItemModel)
            Return
        End If

        ' クラス以外は無視
        If defineKind <> "Class" Then
            Me.InheritsItems = New ObservableCollection(Of InheritsItemModel)
            Return
        End If

        Dim helper = New ReflectionHelper
        Me.InheritsItems = helper.GetInheritsItems(nsView, defineFullName)

    End Sub

    Private Function IsPositionInnerMember(nsView As DataView, ByRef defineFullName As String, ByRef defineKind As String) As Boolean

        ' ※メソッド呼び出し元で、nsView に RowFilter, Sort を適用している状態で、
        ' 渡ってきている可能性があることに注意

        ' 以下のような内側のクラス等の場合、昇順で探すと外側のクラスを採用してしまう
        ' もっとも内側のコンテナを採用したいため、降順で探す（あらかじめ昇順でソートしておく）
        ' Class Class1(0-20)、※開始・終了の定義文字数位置
        '    Class InnerClass(11-17)、※開始・終了の定義文字数位置

        For i As Integer = nsView.Count - 1 To 0 Step -1

            Dim row = nsView(i)
            Dim startLength = CInt(row("StartLength"))
            Dim endLength = CInt(row("EndLength"))

            If startLength <= Me.CaretOffset AndAlso Me.CaretOffset <= endLength Then

                ' 見つけた候補が、Delegate、または Enum の場合、親コンテナに含まれていないかチェック、含まれている場合、その親を候補に取り替える
                defineFullName = CStr(row("DefineFullName"))
                defineKind = CStr(row("DefineKind"))

                If (defineKind = "Delegate") OrElse (defineKind = "Enum") Then

                    ' 入れ子の場合、ドット区切りではなくプラス区切りとなる
                    If defineFullName.Contains("+") Then

                        Dim containerFullName = defineFullName.Substring(0, defineFullName.LastIndexOf("+"))
                        nsView.RowFilter = $"DefineFullName='{containerFullName}' AND DefineKind IN ('Class', 'Structure', 'Interface', 'Module') "
                        nsView.Sort = "StartLength, EndLength"

                        If 0 < nsView.Count Then

                            row = nsView(nsView.Count - 1)
                            defineFullName = CStr(row("DefineFullName"))
                            defineKind = CStr(row("DefineKind"))

                        End If

                    End If

                End If

                Exit For

            End If

        Next

        If String.IsNullOrWhiteSpace(defineFullName) AndAlso String.IsNullOrWhiteSpace(defineKind) Then
            Return False
        Else
            Return True
        End If

    End Function

    Private Sub ShowClassMemberTree()

        Dim nsView = Me.nsViewForMember
        nsView.RowFilter = $"SourceFile='{Me.SourceFile}' AND DefineKind IN ('Class', 'Structure', 'Interface', 'Module', 'Enum', 'Delegate') "
        nsView.Sort = "StartLength, EndLength"

        If nsView.Count = 0 Then
            Me.TreeItems = New ObservableCollection(Of TreeViewItemModel)
            Return
        End If

        Dim defineFullName = String.Empty
        Dim defineKind = String.Empty

        ' クラス定義行の範囲にいるかどうか
        If Not Me.IsPositionInnerMember(nsView, defineFullName, defineKind) Then
            Me.TreeItems = New ObservableCollection(Of TreeViewItemModel)
            Return
        End If

        If defineKind = "Class" Then

            Dim item = Me.GetClassMemberTree(nsView, defineFullName)
            Me.TreeItems = New ObservableCollection(Of TreeViewItemModel) From {item}


        ElseIf defineKind = "Delegate" Then

            nsView.RowFilter = $"DefineFullName='{defineFullName}' AND DefineKind = 'Delegate' "
            If 0 < nsView.Count Then

                Dim headerModel = New TreeViewItemModel With {.Text = "デリゲート", .TreeNodeKind = TreeNodeKinds.DelegateNode, .IsExpanded = True}
                Me.TreeItems = New ObservableCollection(Of TreeViewItemModel) From {headerModel}

                Dim nsRow = nsView(0)
                Dim methodName = CStr(nsRow("DisplayDefineName"))
                Dim returnType = CStr(nsRow("ReturnType"))

                Dim methodArguments = CStr(nsRow("MethodArguments"))
                methodArguments = methodArguments.Substring(methodArguments.IndexOf("("))

                Dim startLength = CInt(nsRow("StartLength"))
                Dim fileName = CStr(nsRow("SourceFile"))

                ' 登録モデルの作成
                Dim childModel = New TreeViewItemModel

                If returnType = "Void" Then
                    childModel.Text = $"{methodName}{methodArguments}"
                Else
                    childModel.Text = $"{methodName}{methodArguments} : {returnType}"
                End If

                childModel.TreeNodeKind = TreeNodeKinds.MethodNode
                childModel.StartLength = startLength
                childModel.FileName = fileName

                headerModel.Children.Add(childModel)

            End If


        ElseIf defineKind = "Enum" Then

            nsView.RowFilter = $"DefineFullName='{defineFullName}' AND DefineKind = 'Enum' "
            If 0 < nsView.Count Then

                Dim headerModel = New TreeViewItemModel With {.Text = "列挙体", .TreeNodeKind = TreeNodeKinds.EnumNode, .IsExpanded = True}
                Me.TreeItems = New ObservableCollection(Of TreeViewItemModel) From {headerModel}

                Dim nsRow = nsView(0)
                Dim defineName = CStr(nsRow("DisplayDefineName"))
                Dim startLength = CInt(nsRow("StartLength"))
                Dim fileName = CStr(nsRow("SourceFile"))
                Dim methodArguments = CStr(nsRow("MethodArguments"))
                Dim memberNames = methodArguments.
                    Split(New String() {","}, StringSplitOptions.RemoveEmptyEntries).
                    Select(Function(x) x.Trim())

                Dim enumModel = New TreeViewItemModel With {.Text = defineName, .TreeNodeKind = TreeNodeKinds.EnumNode, .StartLength = startLength, .FileName = fileName, .IsExpanded = True}
                headerModel.Children.Add(enumModel)

                For Each memberName In memberNames

                    Dim memberModel = New TreeViewItemModel With {.Text = memberName, .TreeNodeKind = TreeNodeKinds.EnumItemNode}
                    enumModel.Children.Add(memberModel)

                Next

            End If


        Else

            Me.TreeItems = New ObservableCollection(Of TreeViewItemModel)

        End If

    End Sub

    Private Function GetClassMemberTree(nsView As DataView, defineFullName As String) As TreeViewItemModel

        ' クラス名
        nsView.RowFilter = $"DefineFullName='{defineFullName}'"
        Dim row = nsView(0)

        Dim parentModel = New TreeViewItemModel
        parentModel.IsExpanded = True
        parentModel.TreeNodeKind = TreeNodeKinds.ClassNode
        parentModel.Text = CStr(row("DisplayDefineName"))
        parentModel.FileName = CStr(row("SourceFile"))
        parentModel.StartLength = CInt(row("StartLength"))
        parentModel.ContainerName = defineFullName

        nsView.RowFilter = $"ContainerFullName='{defineFullName}' AND DefineKind IN ('Field', 'Property', 'Event', 'Constructor', 'Operator', 'Method', 'Delegate', 'Enum') "
        If nsView.Count = 0 Then
            Return parentModel
        End If

        ' クラスメンバー

        ' フィールド
        nsView.RowFilter = $"ContainerFullName='{defineFullName}' AND DefineKind = 'Field' "
        If 0 < nsView.Count Then

            Dim headerModel = New TreeViewItemModel With {.Text = "フィールド", .TreeNodeKind = TreeNodeKinds.FieldNode, .IsExpanded = True}
            parentModel.Children.Add(headerModel)

            For Each nsRow As DataRowView In nsView

                Dim defineName = CStr(nsRow("DisplayDefineName"))
                Dim defineType = CStr(nsRow("DefineType"))
                Dim startLength = CInt(nsRow("StartLength"))
                Dim fileName = CStr(nsRow("SourceFile"))

                Dim childModel = New TreeViewItemModel With {.Text = $"{defineName} : {defineType}", .TreeNodeKind = TreeNodeKinds.FieldNode, .StartLength = startLength, .FileName = fileName}
                headerModel.Children.Add(childModel)

            Next

        End If

        ' プロパティ
        nsView.RowFilter = $"ContainerFullName='{defineFullName}' AND DefineKind = 'Property' "
        If 0 < nsView.Count Then

            Dim headerModel = New TreeViewItemModel With {.Text = "プロパティ", .TreeNodeKind = TreeNodeKinds.PropertyNode, .IsExpanded = True}
            parentModel.Children.Add(headerModel)

            For Each nsRow As DataRowView In nsView

                Dim defineName = CStr(nsRow("DisplayDefineName"))
                Dim defineType = CStr(nsRow("DefineType"))
                Dim startLength = CInt(nsRow("StartLength"))
                Dim fileName = CStr(nsRow("SourceFile"))
                Dim methodArguments = CStr(nsRow("MethodArguments"))

                Dim displayName = $"{defineName} : {defineType}"
                If Not methodArguments.StartsWith("0") Then

                    methodArguments = methodArguments.Substring(methodArguments.IndexOf("("))
                    displayName = $"{defineName}{methodArguments} : {defineType}"

                End If

                Dim childModel = New TreeViewItemModel With {.Text = displayName, .TreeNodeKind = TreeNodeKinds.PropertyNode, .StartLength = startLength, .FileName = fileName}
                headerModel.Children.Add(childModel)

            Next

        End If

        ' イベント定義
        nsView.RowFilter = $"ContainerFullName='{defineFullName}' AND DefineKind = 'Event' "
        If 0 < nsView.Count Then

            Dim headerModel = New TreeViewItemModel With {.Text = "イベント定義", .TreeNodeKind = TreeNodeKinds.EventNode, .IsExpanded = True}
            parentModel.Children.Add(headerModel)

            For Each nsRow As DataRowView In nsView

                Dim defineName = CStr(nsRow("DisplayDefineName"))
                Dim defineType = CStr(nsRow("DefineType"))
                Dim startLength = CInt(nsRow("StartLength"))
                Dim fileName = CStr(nsRow("SourceFile"))

                Dim childModel = New TreeViewItemModel

                If defineType = String.Empty Then
                    childModel.Text = $"{defineName}"
                Else
                    childModel.Text = $"{defineName} : {defineType}"
                End If

                childModel.TreeNodeKind = TreeNodeKinds.EventNode
                childModel.StartLength = startLength
                childModel.FileName = fileName

                headerModel.Children.Add(childModel)

            Next

        End If

        ' コンストラクタ
        nsView.RowFilter = $"ContainerFullName='{defineFullName}' AND DefineKind = 'Constructor' "
        If 0 < nsView.Count Then

            Dim headerModel = New TreeViewItemModel With {.Text = "コンストラクタ", .TreeNodeKind = TreeNodeKinds.MethodNode, .IsExpanded = True}
            parentModel.Children.Add(headerModel)

            For Each nsRow As DataRowView In nsView

                Dim methodName = CStr(nsRow("DisplayDefineName"))
                Dim returnType = CStr(nsRow("ReturnType"))

                Dim methodArguments = CStr(nsRow("MethodArguments"))
                methodArguments = methodArguments.Substring(methodArguments.IndexOf("("))

                Dim startLength = CInt(nsRow("StartLength"))
                Dim fileName = CStr(nsRow("SourceFile"))

                ' 登録モデルの作成
                Dim childModel = New TreeViewItemModel

                If returnType = "Void" Then
                    childModel.Text = $"{methodName}{methodArguments}"
                Else
                    childModel.Text = $"{methodName}{methodArguments} : {returnType}"
                End If

                childModel.TreeNodeKind = TreeNodeKinds.MethodNode
                childModel.StartLength = startLength
                childModel.FileName = fileName

                headerModel.Children.Add(childModel)

            Next

        End If

        ' Windows API
        nsView.RowFilter = $"ContainerFullName='{defineFullName}' AND DefineKind = 'WindowsAPI' "
        If 0 < nsView.Count Then

            Dim headerModel = New TreeViewItemModel With {.Text = "Windows API", .TreeNodeKind = TreeNodeKinds.MethodNode, .IsExpanded = True}
            parentModel.Children.Add(headerModel)

            For Each nsRow As DataRowView In nsView

                Dim methodName = CStr(nsRow("DisplayDefineName"))
                Dim returnType = CStr(nsRow("ReturnType"))

                Dim methodArguments = CStr(nsRow("MethodArguments"))
                methodArguments = methodArguments.Substring(methodArguments.IndexOf("("))

                Dim startLength = CInt(nsRow("StartLength"))
                Dim fileName = CStr(nsRow("SourceFile"))

                ' 登録モデルの作成
                Dim childModel = New TreeViewItemModel

                If returnType = "Void" Then
                    childModel.Text = $"{methodName}{methodArguments}"
                Else
                    childModel.Text = $"{methodName}{methodArguments} : {returnType}"
                End If

                childModel.TreeNodeKind = TreeNodeKinds.MethodNode
                childModel.StartLength = startLength
                childModel.FileName = fileName

                headerModel.Children.Add(childModel)

            Next

        End If

        ' オペレーター
        nsView.RowFilter = $"ContainerFullName='{defineFullName}' AND DefineKind = 'Operator' "
        If 0 < nsView.Count Then

            Dim headerModel = New TreeViewItemModel With {.Text = "オペレーター", .TreeNodeKind = TreeNodeKinds.OperatorNode, .IsExpanded = True}
            parentModel.Children.Add(headerModel)

            For Each nsRow As DataRowView In nsView

                Dim methodName = CStr(nsRow("DisplayDefineName"))
                Dim returnType = CStr(nsRow("ReturnType"))

                Dim methodArguments = CStr(nsRow("MethodArguments"))
                methodArguments = methodArguments.Substring(methodArguments.IndexOf("("))

                Dim startLength = CInt(nsRow("StartLength"))
                Dim fileName = CStr(nsRow("SourceFile"))

                ' 登録モデルの作成
                Dim childModel = New TreeViewItemModel

                If returnType = "Void" Then
                    childModel.Text = $"{methodName}{methodArguments}"
                Else
                    childModel.Text = $"{methodName}{methodArguments} : {returnType}"
                End If

                childModel.TreeNodeKind = TreeNodeKinds.MethodNode
                childModel.StartLength = startLength
                childModel.FileName = fileName

                headerModel.Children.Add(childModel)

            Next

        End If

        ' イベントハンドラ、メソッド
        nsView.RowFilter = $"ContainerFullName='{defineFullName}' AND DefineKind = 'Method' "
        If 0 < nsView.Count Then

            Dim eventHandlerItems = New List(Of TreeViewItemModel)
            Dim methodItems = New List(Of TreeViewItemModel)

            For Each nsRow As DataRowView In nsView

                Dim methodName = CStr(nsRow("DisplayDefineName"))
                Dim returnType = CStr(nsRow("ReturnType"))

                Dim methodArguments = CStr(nsRow("MethodArguments"))
                Dim argumentCount = CInt(methodArguments.Substring(0, methodArguments.IndexOf("(")))
                methodArguments = methodArguments.Substring(methodArguments.IndexOf("("))

                Dim startLength = CInt(nsRow("StartLength"))
                Dim fileName = CStr(nsRow("SourceFile"))

                ' 登録モデルの作成
                Dim childModel = New TreeViewItemModel

                If returnType = "Void" Then
                    childModel.Text = $"{methodName}{methodArguments}"
                Else
                    childModel.Text = $"{methodName}{methodArguments} : {returnType}"
                End If

                childModel.TreeNodeKind = TreeNodeKinds.MethodNode
                childModel.StartLength = startLength
                childModel.FileName = fileName

                ' 引数が２つで、１つ目が Object 型、２つ目の型の最後が EventArgs で終わっている場合、イベントハンドラと判断する
                Dim isEventHandler = False
                If argumentCount = 2 Then

                    Dim argument = methodArguments.Substring(1)
                    argument = argument.Substring(0, argument.Length - 1)
                    Dim arguments = argument.Split(New String() {","}, StringSplitOptions.RemoveEmptyEntries)

                    If arguments(0).ToLower().EndsWith("object") AndAlso arguments(1).ToLower().EndsWith("eventargs") Then
                        isEventHandler = True
                    End If

                End If

                If isEventHandler Then
                    ' EventHandler　
                    eventHandlerItems.Add(childModel)
                Else
                    ' Method
                    methodItems.Add(childModel)
                End If

            Next

            ' イベントハンドラ
            If 0 < eventHandlerItems.Count Then

                Dim headerModel = New TreeViewItemModel With {.Text = "イベントハンドラ", .TreeNodeKind = TreeNodeKinds.MethodNode, .IsExpanded = True}
                parentModel.Children.Add(headerModel)
                eventHandlerItems.ForEach(Sub(x) headerModel.Children.Add(x))

            End If

            ' メソッド
            If 0 < methodItems.Count Then

                Dim headerModel = New TreeViewItemModel With {.Text = "メソッド", .TreeNodeKind = TreeNodeKinds.MethodNode, .IsExpanded = True}
                parentModel.Children.Add(headerModel)
                methodItems.ForEach(Sub(x) headerModel.Children.Add(x))

            End If

        End If

        ' デリゲート
        nsView.RowFilter = $"ContainerFullName='{defineFullName}' AND DefineKind = 'Delegate' "
        If 0 < nsView.Count Then

            Dim headerModel = New TreeViewItemModel With {.Text = "デリゲート", .TreeNodeKind = TreeNodeKinds.DelegateNode, .IsExpanded = True}
            parentModel.Children.Add(headerModel)

            For Each nsRow As DataRowView In nsView

                Dim methodName = CStr(nsRow("DisplayDefineName"))
                Dim returnType = CStr(nsRow("ReturnType"))

                Dim methodArguments = CStr(nsRow("MethodArguments"))
                methodArguments = methodArguments.Substring(methodArguments.IndexOf("("))

                Dim startLength = CInt(nsRow("StartLength"))
                Dim fileName = CStr(nsRow("SourceFile"))

                ' 登録モデルの作成
                Dim childModel = New TreeViewItemModel

                If returnType = "Void" Then
                    childModel.Text = $"{methodName}{methodArguments}"
                Else
                    childModel.Text = $"{methodName}{methodArguments} : {returnType}"
                End If

                childModel.TreeNodeKind = TreeNodeKinds.MethodNode
                childModel.StartLength = startLength
                childModel.FileName = fileName

                headerModel.Children.Add(childModel)

            Next

        End If

        ' 列挙体
        nsView.RowFilter = $"ContainerFullName='{defineFullName}' AND DefineKind = 'Enum' "
        If 0 < nsView.Count Then

            Dim headerModel = New TreeViewItemModel With {.Text = "列挙体", .TreeNodeKind = TreeNodeKinds.EnumNode, .IsExpanded = True}
            parentModel.Children.Add(headerModel)

            For Each nsRow As DataRowView In nsView

                Dim defineName = CStr(nsRow("DisplayDefineName"))
                Dim startLength = CInt(nsRow("StartLength"))
                Dim fileName = CStr(nsRow("SourceFile"))
                Dim methodArguments = CStr(nsRow("MethodArguments"))
                Dim memberNames = methodArguments.
                    Split(New String() {","}, StringSplitOptions.RemoveEmptyEntries).
                    Select(Function(x) x.Trim())

                Dim enumModel = New TreeViewItemModel With {.Text = defineName, .TreeNodeKind = TreeNodeKinds.EnumNode, .StartLength = startLength, .FileName = fileName, .IsExpanded = True}
                headerModel.Children.Add(enumModel)

                For Each memberName In memberNames

                    Dim memberModel = New TreeViewItemModel With {.Text = memberName, .TreeNodeKind = TreeNodeKinds.EnumItemNode}
                    enumModel.Children.Add(memberModel)

                Next

            Next

        End If

        Return parentModel

    End Function



    ' 今まで Codimension を意識して作成したと記載していて、これからもその通りなのだが、
    ' MVVM 的な実現技術は、ufcpp 様のサンプルを参考にさせていただきました。

    ' [サンプル] 式木を WPF で GUI 表示
    ' http://ufcpp.net/study/csharp/sm_treeview.html
    ' → 「ソース一式（ZIP 形式）」のリンクからダウンロード

    ' Content プロパティに ViewModel をバインドさせて、後は DataTemplate に任せる方法
    ' ResourceDictionary に表示形式の設定をまとめて用意していて、こちらを参照している
    ' → クラスごとに DataTemplate を定義していて、クラスごとに違うレイアウトを、再帰して表示させる方法

    Private Sub ShowMethodFlowchart()

        ' 今現在のキャレット位置が、任意のメソッド定義位置の範囲にいるかチェック
        Dim nsView = Me.nsViewForMethod
        nsView.RowFilter = $"SourceFile='{Me.SourceFile}' AND DefineKind IN ('Constructor', 'Operator', 'WindowsAPI', 'EventHandler', 'Method') "
        nsView.Sort = "StartLength, EndLength"

        If nsView.Count = 0 Then
            Me.MethodModel = New MethodTemplateModel With {.Signature = "Clear"}
            Return
        End If

        ' DataView.RowFilter は、文字列の条件分岐がある場合、追加で数値の比較を書いても、うまくフィルタされない？ため別判定
        Dim methodRange = String.Empty
        For Each row As DataRowView In nsView

            Dim startLength = CInt(row("StartLength"))
            Dim endLength = CInt(row("EndLength"))

            ' メソッドの範囲内なら、ソースコードからメソッドの定義範囲の文字列を取得
            If startLength <= Me.CaretOffset AndAlso Me.CaretOffset <= endLength Then
                methodRange = Me.SourceCode.Substring(startLength, endLength - startLength)
                Exit For
            End If

        Next

        If methodRange = String.Empty Then
            Me.MethodModel = New MethodTemplateModel With {.Signature = "Clear"}
            Return
        End If

        Dim walker = New MethodSyntaxWalker
        walker.Parse(methodRange)
        Me.MethodModel = walker.MethodModel

    End Sub

#End Region

#Region "メンバーツリーのノード選択"
    Private _SelectedItemChangedCommand As ListenerCommand(Of TreeViewItemModel)

    Public ReadOnly Property SelectedItemChangedCommand() As ListenerCommand(Of TreeViewItemModel)
        Get
            If _SelectedItemChangedCommand Is Nothing Then
                _SelectedItemChangedCommand = New ListenerCommand(Of TreeViewItemModel)(AddressOf SelectedItemChanged)
            End If
            Return _SelectedItemChangedCommand
        End Get
    End Property

    Private Sub SelectedItemChanged(ByVal e As TreeViewItemModel)

        If (e Is Nothing) OrElse (e.FileName <> Me.SourceFile) Then
            Return
        End If

        Me.CaretOffset = e.StartLength

    End Sub


    ' 不具合
    ' LivetCallMethodAction 経由（メソッド直接バインディング）だと、NullReferenceException が発生してしまう

    ' 現象再現手順
    ' エディタ内、クラス範囲内の任意の位置をクリックしてキャレット位置を更新 → その後クラスメンバーツリーで任意のメソッドノードをクリック → その後エディタ内をクリックすると、NullReferenceException

    ' Livet.dll 内
    ' System.NullReferenceException はハンドルされませんでした。
    ' HResult=-2147467261
    ' Message=オブジェクト参照がオブジェクト インスタンスに設定されていません。
    ' 場所 Livet.Behaviors.MethodBinderWithArgument.Invoke(Object targetObject, String methodName, Object argument)

    ' 例外エラーの発生場所が、TreeItems のセッター内、RaisePropertyChanged であり、ソースを確認したが分からず。
    ' https://github.com/ugaya40/Livet/blob/master/.NET4.0/Livet(.NET4.0)/Behaviors/MethodBinderWithArgument.cs

    ' 対策１
    ' ObservableCollection 内のデータの扱い方について、都度インスタンス生成し直しする扱い方がまずいのかと思い、
    ' Clear メソッドと Add メソッドで書き換えてみたが、それでも現象が出た

    ' 対策２
    ' メソッド直接バインディングを止めて、コマンドバインディングに変更することで回避可能だった（MethodBinderWithArgument.Invoke メソッドを通らない実行経路に変えた）
    ' WPF: TreeViewItem bound to an ICommand
    ' https://stackoverflow.com/questions/2266890/wpf-treeviewitem-bound-to-an-icommand



    'Public Sub MemberTree_SelectedItemChanged(e As TreeViewItemModel)
    '    Console.WriteLine(e)
    'End Sub

#End Region

End Class
